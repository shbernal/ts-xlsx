// Adapter: binds the corpus's implementation-blind contract vocabulary to the
// *rewrite* — the new, strict-TypeScript library under src/ (Phase 3). This is the
// sibling of current.mjs: the same cases run unchanged against it, and the finish
// line for an area is every one of its baselines flipping to `pass` here.
//
// Node 24 runs the .ts sources directly (type-stripping), so this adapter imports
// them with no build step. Strict type-checking is enforced separately by
// `npm run typecheck` (tsc --noEmit -p tsconfig.build.json).
//
// The rewrite is incomplete by construction: it grows one module at a time. Any
// capability it does not yet implement is served by a tagged thrower (see the
// Proxy below) so the runner SKIPS the cases that need it instead of reporting
// false regressions. As a module lands, its capability moves into `impl` and the
// corresponding cases light up and must go green.
//
// Feature-gating: the writer covers only part of the spec vocabulary so far
// (worksheets; number/string/boolean/formula cells; the four core properties). A spec
// that reaches for anything else is served the SAME `notImplemented` skip, so a
// partially-built writer never produces a false regression — it only runs the cases it
// can faithfully serialize, and those must go green.

import {createRequire} from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import {Duplex, PassThrough} from 'node:stream';
import {tmpdir} from 'node:os';
import {fileURLToPath} from 'node:url';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {decodeAddress, decodeRange, encodeAddress} from '../../../src/core/address.ts';
import {detectValueType} from '../../../src/core/value.ts';
import {Workbook} from '../../../src/core/workbook.ts';
import {readXlsx} from '../../../src/io/xlsx/read.ts';
import {readWorkbookStream} from '../../../src/io/xlsx/read-rows.ts';
import {writeXlsx} from '../../../src/io/xlsx/write.ts';
import {readCsv} from '../../../src/io/csv/read.ts';
import {writeCsv, writeCsvText} from '../../../src/io/csv/write.ts';
import {WorkbookStreamWriter} from '../../../src/io/xlsx/write-stream.ts';
import {packageFacts} from './ooxml-facts.mjs';

// JSZip is an independent zip implementation used only to VERIFY the streaming writer's output (CRC
// integrity), a hostile-input posture toward our own archive — never in the production src path.
const require = createRequire(import.meta.url);
const JSZip = require('jszip');

// Durable sample inputs live under test/corpus/fixtures/<case-slug>/ — the SAME tree the
// oracle adapter reads, so a fixture-backed case measures both implementations against one
// real-world file. The rewrite reads them straight through readXlsx (a fixture is just a
// foreign `.xlsx` buffer).
const FIXTURES_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'fixtures');
const fixtureBytes = rel => fs.readFileSync(path.join(FIXTURES_ROOT, rel));
const readFixture = rel => readXlsx(fixtureBytes(rel));

// The 1-based `row.values` array a full-load reader exposes, rebuilt from a streamed row's cells:
// index 0 is an empty leading slot and column A lands at index 1, so streaming and buffered reads
// index identically. Gaps (and the leading slot) are null, every present value normalized.
const streamedRowValues = cells => {
  const width = cells.reduce((max, cell) => Math.max(max, cell.col), 0);
  const values = new Array(width + 1).fill(null);
  for (const cell of cells) values[cell.col] = normalizeStreamValue(cell.value);
  return values;
};

// A 1×1 PNG — a minimal image payload for anchoring on a sheet.
const ONE_PX_PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64'
  )
);

const notImplemented = message => {
  const err = new Error(`rewrite: ${message}`);
  err.notImplemented = true;
  return err;
};

const SUPPORTED_TOP_KEYS = new Set(['sheets', 'properties', 'definedNames']);
const SUPPORTED_PROP_KEYS = new Set(['creator', 'lastModifiedBy', 'created', 'modified']);
const SUPPORTED_SHEET_KEYS = new Set([
  'name', 'state', 'cells', 'columns', 'rows', 'properties', 'pageSetup', 'pageMargins', 'headerFooter', 'tables', 'merges', 'autoFilter', 'images', 'background',
]);
const SUPPORTED_CELL_KEYS = new Set(['ref', 'value', 'formula', 'sharedFormula', 'result', 'hyperlink', 'text', 'tooltip', 'fill', 'numFmt', 'font', 'border', 'alignment', 'protection', 'note']);
const SUPPORTED_SHEET_PROP_KEYS = new Set(['defaultRowHeight', 'defaultColWidth']);
const SUPPORTED_COLUMN_KEYS = new Set(['index', 'width', 'hidden', 'numFmt', 'fill', 'font', 'border', 'alignment', 'protection']);
const SUPPORTED_ROW_KEYS = new Set(['index', 'height', 'hidden', 'outlineLevel', 'collapsed', 'fill']);
const SUPPORTED_PAGE_MARGIN_KEYS = new Set(['left', 'right', 'top', 'bottom', 'header', 'footer']);
const SUPPORTED_PAGE_SETUP_KEYS = new Set(['fitToPage', 'fitToWidth', 'fitToHeight', 'scale', 'orientation', 'pageOrder', 'paperSize']);
const SUPPORTED_HEADER_FOOTER_KEYS = new Set([
  'oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter',
]);
const SUPPORTED_TABLE_KEYS = new Set([
  'name', 'ref', 'headers', 'columnDefs', 'rows', 'headerRow', 'totalsRow',
]);
const SUPPORTED_TABLE_COLUMN_KEYS = new Set(['name', 'totalsRowLabel', 'totalsRowFunction']);

// Build an _xlnm.Print_Area refersTo from a comma-separated area (e.g. 'A1:F10,A12:F21' or 'A:D'):
// each range is made absolute ($-prefixed on every column and row bound) and sheet-qualified, exactly
// how Excel records a print area. A whole-column range keeps its column-only shape ($A:$D).
const absolutizeRef = ref => ref.replace(/([A-Z]+)/g, '$$$1').replace(/(\d+)/g, '$$$1');
const printAreaRefersTo = (sheetName, area) =>
  area
    .split(',')
    .map(range => `${sheetName}!${absolutizeRef(range)}`)
    .join(',');

const toDate = v => (v && typeof v === 'object' && v.invalidDate ? new Date(NaN) : new Date(v));
const isoOrNull = d => (d instanceof Date && !Number.isNaN(d.getTime()) ? d.toISOString() : null);

// A JSON-serializable view of a read-back cell value: a Date becomes { date: iso } (null when
// invalid), every other object is deep-cloned, and a scalar passes through. Mirrors the oracle.
const normalizeStreamValue = v => {
  if (v instanceof Date) return {date: Number.isNaN(v.getTime()) ? null : v.toISOString()};
  if (v && typeof v === 'object') return JSON.parse(JSON.stringify(v));
  return v ?? null;
};

// Some specs express a rich-text run in the flat inline shape `{ text, bold, italic, … }`, while the
// rewrite models a run as `{ text, font: { … } }`. Translate a spec value into the model shape on the
// way in…
const specValueToModel = value => {
  if (value && typeof value === 'object' && Array.isArray(value.richText)) {
    return {
      richText: value.richText.map(run => {
        const {text, ...font} = run;
        return Object.keys(font).length ? {text, font} : {text};
      }),
    };
  }
  return value;
};

// …and flatten a read-back run's `font` facets back onto the run on the way out, so a spec asserting
// on `run.bold` sees the shape it wrote.
const modelValueToSpec = value => {
  if (value && typeof value === 'object' && Array.isArray(value.richText)) {
    return {richText: value.richText.map(({text, font}) => ({text, ...(font || {})}))};
  }
  return value;
};

// Map a declarative spec onto the rewrite's Workbook model, throwing a `notImplemented`
// skip the moment the spec uses a feature the writer cannot represent yet.
function buildFrom(spec = {}) {
  for (const k of Object.keys(spec)) {
    if (!SUPPORTED_TOP_KEYS.has(k)) throw notImplemented(`spec.${k} not supported yet`);
  }
  const workbook = new Workbook();

  const p = spec.properties || {};
  for (const k of Object.keys(p)) {
    if (!SUPPORTED_PROP_KEYS.has(k)) throw notImplemented(`properties.${k} not supported yet`);
  }
  if (p.creator !== undefined) workbook.properties.creator = p.creator;
  if (p.lastModifiedBy !== undefined) workbook.properties.lastModifiedBy = p.lastModifiedBy;
  if (p.created !== undefined) workbook.properties.created = toDate(p.created);
  if (p.modified !== undefined) workbook.properties.modified = toDate(p.modified);

  for (const s of spec.sheets || []) {
    for (const k of Object.keys(s)) {
      if (!SUPPORTED_SHEET_KEYS.has(k)) throw notImplemented(`sheet.${k} not supported yet`);
    }
    const sheet = workbook.addWorksheet(s.name, s.state ? {state: s.state} : undefined);

    const sp = s.properties || {};
    for (const k of Object.keys(sp)) {
      if (!SUPPORTED_SHEET_PROP_KEYS.has(k)) throw notImplemented(`sheet.properties.${k} not supported yet`);
    }
    if (sp.defaultRowHeight !== undefined) sheet.properties.defaultRowHeight = sp.defaultRowHeight;
    if (sp.defaultColWidth !== undefined) sheet.properties.defaultColWidth = sp.defaultColWidth;

    const pm = s.pageMargins || {};
    for (const k of Object.keys(pm)) {
      if (!SUPPORTED_PAGE_MARGIN_KEYS.has(k)) throw notImplemented(`pageMargins.${k} not supported yet`);
      sheet.pageMargins[k] = pm[k];
    }

    const psu = s.pageSetup || {};
    for (const k of Object.keys(psu)) {
      if (!SUPPORTED_PAGE_SETUP_KEYS.has(k)) throw notImplemented(`pageSetup.${k} not supported yet`);
      sheet.pageSetup[k] = psu[k];
    }

    const hf = s.headerFooter || {};
    for (const k of Object.keys(hf)) {
      if (!SUPPORTED_HEADER_FOOTER_KEYS.has(k)) throw notImplemented(`headerFooter.${k} not supported yet`);
      sheet.headerFooter[k] = hf[k];
    }

    for (const t of s.tables || []) {
      for (const k of Object.keys(t)) {
        if (!SUPPORTED_TABLE_KEYS.has(k)) throw notImplemented(`table.${k} not supported yet`);
      }
      let columns;
      if (t.columnDefs) {
        for (const cd of t.columnDefs) {
          for (const k of Object.keys(cd)) {
            if (!SUPPORTED_TABLE_COLUMN_KEYS.has(k)) throw notImplemented(`table.columnDefs.${k} not supported yet`);
          }
        }
        columns = t.columnDefs.map(cd => {
          const col = {name: cd.name};
          if (cd.totalsRowLabel !== undefined) col.totalsRowLabel = cd.totalsRowLabel;
          if (cd.totalsRowFunction !== undefined) col.totalsRowFunction = cd.totalsRowFunction;
          return col;
        });
      } else {
        columns = (t.headers || []).map(name => ({name}));
      }
      // A spec may express a table ref as the full occupied range (`A1:B3`), the shape the oracle
      // accepts, while the model anchors at the single top-left cell and derives the range from the
      // row count. Take the anchor; the declared row count reconstructs the same range.
      const options = {name: t.name, ref: t.ref.split(':')[0], columns, rowCount: (t.rows || []).length};
      if (t.headerRow !== undefined) options.headerRow = t.headerRow;
      if (t.totalsRow !== undefined) options.totalsRow = t.totalsRow;
      sheet.addTable(options);
    }

    for (const range of s.merges || []) sheet.mergeCells(range);

    // A spec autoFilter is either a bare range string or the structured {ref, columns} shape; the
    // model's setter accepts both, so pass it through verbatim.
    if (s.autoFilter !== undefined) sheet.autoFilter = s.autoFilter;

    for (const col of s.columns || []) {
      for (const k of Object.keys(col)) {
        if (!SUPPORTED_COLUMN_KEYS.has(k)) throw notImplemented(`column.${k} not supported yet`);
      }
      const target = sheet.getColumn(col.index);
      if (col.width !== undefined) target.width = col.width;
      if (col.hidden !== undefined) target.hidden = col.hidden;
      if (col.numFmt !== undefined) target.numFmt = col.numFmt;
      if (col.fill !== undefined) target.fill = col.fill;
      if (col.font !== undefined) target.font = col.font;
      if (col.border !== undefined) target.border = col.border;
      if (col.alignment !== undefined) target.alignment = col.alignment;
      if (col.protection !== undefined) target.protection = col.protection;
    }

    for (const row of s.rows || []) {
      for (const k of Object.keys(row)) {
        if (!SUPPORTED_ROW_KEYS.has(k)) throw notImplemented(`row.${k} not supported yet`);
      }
      const target = sheet.getRow(row.index);
      if (row.height !== undefined) target.height = row.height;
      if (row.hidden !== undefined) target.hidden = row.hidden;
      if (row.outlineLevel !== undefined) target.outlineLevel = row.outlineLevel;
      if (row.collapsed !== undefined) target.collapsed = row.collapsed;
      if (row.fill !== undefined) target.fill = row.fill;
    }

    // Images build after columns and rows so a fractional anchor resolves against the sheet's real
    // column widths and row heights (the model floors + offsets at addImage time).
    for (const img of s.images || []) {
      // A spec omits `extension` to mean the default 'png'; it sets the key (to a dirty or missing
      // value) on purpose to exercise the library's write-side extension sanitisation. Pass the raw
      // value through — `workbook.addImage` normalises a leading dot / query string / missing hint.
      const options = 'extension' in img ? {buffer: ONE_PX_PNG, extension: img.extension} : {buffer: ONE_PX_PNG};
      anchorSpecImage(sheet, workbook.addImage(options), img.range);
    }
    // A sheet background is a workbook image tiled behind the grid, not anchored — it rides its own
    // worksheet `<picture>` relationship, so a case can assert it coexists with comment/VML parts.
    if (s.background) {
      sheet.addBackgroundImage(workbook.addImage({buffer: ONE_PX_PNG, extension: s.background.extension || 'png'}));
    }

    for (const c of s.cells || []) {
      for (const k of Object.keys(c)) {
        if (!SUPPORTED_CELL_KEYS.has(k)) throw notImplemented(`cell.${k} not supported yet`);
      }
      const cell = sheet.getCell(c.ref);
      if ('hyperlink' in c) {
        // The display label is a plain string or a rich-text value; both serialise faithfully.
        cell.value = {
          hyperlink: c.hyperlink,
          text: c.text ?? '',
          ...(c.tooltip !== undefined ? {tooltip: c.tooltip} : {}),
        };
      } else if ('formula' in c) {
        cell.value = 'result' in c ? {formula: c.formula, result: c.result} : {formula: c.formula};
      } else if ('sharedFormula' in c) {
        // A shared-formula clone names its master by address; the master is a plain formula cell.
        cell.value =
          'result' in c ? {sharedFormula: c.sharedFormula, result: c.result} : {sharedFormula: c.sharedFormula};
      } else if ('value' in c) {
        const v = c.value;
        if (v !== null && typeof v === 'object') {
          // A structured date value materializes a Date; every other object shape is a value
          // kind the writer does not model yet, so skip the behavior rather than mis-serialize.
          if (v.invalidDate) cell.value = new Date(NaN);
          else if (v.date) cell.value = toDate(v.date);
          else if (Array.isArray(v.richText)) cell.value = {richText: v.richText};
          else throw notImplemented(`cell value shape ${JSON.stringify(v)} not supported yet`);
        } else {
          cell.value = v;
        }
      }
      if (c.fill !== undefined) cell.fill = c.fill;
      if (c.numFmt !== undefined) cell.numFmt = c.numFmt;
      if (c.font !== undefined) cell.font = c.font;
      if (c.border !== undefined) cell.border = c.border;
      if (c.alignment !== undefined) cell.alignment = c.alignment;
      if (c.protection !== undefined) cell.protection = c.protection;
      // A note attaches a comments part + legacy VML drawing; a case pairs it with a background image
      // to assert the two features' worksheet relationships never collide.
      if (c.note !== undefined) cell.note = c.note;
    }
  }

  // Workbook-level defined names are added after every sheet exists, since a scoped name targets a
  // sheet by name. The corpus spec expresses a name as one-or-more ranges; the model stores a single
  // refersTo formula, so the ranges join into the comma-separated union OOXML persists in one element.
  for (const dn of spec.definedNames || []) {
    workbook.defineName({
      name: dn.name,
      refersTo: (dn.ranges || []).join(','),
      ...(dn.scope !== undefined ? {scope: dn.scope} : {}),
    });
  }
  return workbook;
}

// Mirror current.mjs's normalizeCell for the rewrite's Cell: a plain JSON view of the
// value that survived the round-trip. Style facets are absent until the reader reads
// them, matching the contract that an unmaterialized facet is simply not present.
function normalizeRewriteCell(cell) {
  const v = cell.value;
  let out;
  if (v && typeof v === 'object' && 'hyperlink' in v) {
    out = {hyperlink: v.hyperlink, text: v.text, ...(v.tooltip !== undefined ? {tooltip: v.tooltip} : {})};
  } else if (v && typeof v === 'object' && 'sharedFormula' in v) {
    out = {sharedFormula: v.sharedFormula, formula: v.formula ?? null, result: v.result ?? null};
  } else if (v && typeof v === 'object' && 'formula' in v) out = {formula: v.formula, result: v.result ?? null};
  else if (v instanceof Date) out = {value: Number.isNaN(v.getTime()) ? null : v.toISOString()};
  else out = {value: v ?? null};
  // A style facet is reported only when the round-trip materialized it, matching the
  // contract that an unset facet is simply absent (never an empty placeholder).
  if (cell.fill !== undefined) out.fill = cell.fill;
  if (cell.numFmt) out.numFmt = cell.numFmt;
  if (cell.font !== undefined) out.font = cell.font;
  if (cell.border !== undefined) out.border = cell.border;
  if (cell.alignment !== undefined) out.alignment = cell.alignment;
  if (cell.protection !== undefined) out.protection = cell.protection;
  // A note is cell metadata, reported only when the round-trip preserved one — mirrors the oracle so a
  // case can assert a comment survives alongside a table/background rather than reading undefined.
  if (cell.note !== undefined) out.note = cell.note;
  return out;
}

// Decompose an ExcelJS-shaped aggregate style object onto the rewrite's per-facet setters.
// The rewrite has no `.style` aggregate: each cell owns independent facet fields, so "assign
// one base style to two cells" (the shared-style aliasing setup) is just assigning each facet
// present. Assigning the SAME base object to two cells shares the facet references — exactly the
// aliasing a copy-on-write setter must not let bleed when one cell is later mutated.
function applyStyle(cell, style) {
  if (style.fill !== undefined) cell.fill = style.fill;
  if (style.numFmt !== undefined) cell.numFmt = style.numFmt;
  if (style.font !== undefined) cell.font = style.font;
  if (style.border !== undefined) cell.border = style.border;
  if (style.alignment !== undefined) cell.alignment = style.alignment;
  if (style.protection !== undefined) cell.protection = style.protection;
}

function partMapOf(buffer) {
  const unzipped = unzipSync(buffer);
  const out = {};
  for (const name of Object.keys(unzipped)) out[name] = strFromU8(unzipped[name]);
  return out;
}

// Package-part facts a passthrough round-trip must preserve — the mirror of the oracle's
// `packageFactsFromZip`: counts of part families the reader does not fully model (drawings, VML,
// media, pivot tables/caches, comments) plus the worksheet/drawing reference flags that wire
// unmodeled features (a vector-shape drawing, a header/footer image) into the sheet.
const packagePartFacts = parts => {
  const names = Object.keys(parts);
  const at = rx => parts[names.find(p => rx.test(p))] ?? '';
  const ws1 = at(/worksheets\/sheet1\.xml$/);
  const drawing1 = at(/drawings\/drawing1\.xml$/);
  return {
    drawings: names.filter(p => /xl\/drawings\/drawing\d+\.xml$/.test(p)).length,
    vml: names.filter(p => /vmlDrawing\d+\.vml$/.test(p)).length,
    media: names.filter(p => /xl\/media\//.test(p)).length,
    pivotTables: names.filter(p => /pivotTables\/pivotTable\d+\.xml$/.test(p)).length,
    pivotCache: names.filter(p => /pivotCache\/.+\.xml$/.test(p)).length,
    slicers: names.filter(p => /slicer/i.test(p)).length,
    comments: names.filter(p => /comments\d+\.xml$/.test(p)).length,
    hasLegacyDrawingHF: /<legacyDrawingHF\b/.test(ws1),
    hasDrawingRef: /<drawing\b/.test(ws1),
    hasHeaderFooterImageToken: /&amp;G|&G/.test(ws1),
    drawingHasShape: /<xdr:sp\b/.test(drawing1),
    drawingHasPicture: /<xdr:pic\b/.test(drawing1),
  };
};

const hexBytes = hex => Uint8Array.from(hex.match(/../g).map(h => parseInt(h, 16)));

// Translate a corpus image range — a string like "B2:D6", or a {tl, br?/ext?, editAs?} object — into
// the model's typed addImage call. A one-cell anchor is a point plus a fixed pixel extent (editAs is a
// two-cell-only attribute the model drops by construction); a two-cell anchor spans tl..br. A
// fractional grid coordinate (col 3.5) is passed through — the model floors it to the cell and derives
// the sub-cell EMU offset from that cell's real width/height.
function anchorSpecImage(sheet, imageId, range) {
  if (typeof range === 'string') {
    const {left, top, right, bottom} = decodeRange(range);
    sheet.addImage(imageId, {tl: {col: left - 1, row: top - 1}, br: {col: right, row: bottom}});
    return;
  }
  const {tl, br, ext, editAs} = range;
  if (ext !== undefined) {
    sheet.addImage(imageId, {tl, ext: {width: ext.width, height: ext.height}});
  } else {
    sheet.addImage(imageId, editAs !== undefined ? {tl, br, editAs} : {tl, br});
  }
}

// Parse the integer children of an <xdr:from>/<xdr:to> block, mirroring the oracle so a drawing anchor
// reports the same plain-number geometry from either adapter.
const intAt = (xml, tag) => {
  const m = xml.match(new RegExp(`<${tag}>(-?\\d+)</${tag}>`));
  return m ? Number(m[1]) : null;
};
const parseAnchorSide = block =>
  block
    ? {col: intAt(block, 'xdr:col'), colOff: intAt(block, 'xdr:colOff'), row: intAt(block, 'xdr:row'), rowOff: intAt(block, 'xdr:rowOff')}
    : null;
const imageXmlWellFormed = xml => !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml);

// Expand an OOXML sqref (space-separated ranges) into its covered cell references, bounded by a cap so
// a whole-column range never balloons — used to check that a range-form validation is reported on
// every covered cell. An unbounded whole-row/column part is skipped rather than expanded.
function expandSqref(sqref, cap = 4096) {
  const refs = [];
  for (const part of String(sqref).split(/\s+/).filter(Boolean)) {
    const {left, right, top, bottom} = decodeRange(part);
    if (left == null || right == null || top == null || bottom == null) continue;
    for (let c = left; c <= right && refs.length < cap; c++) {
      for (let r = top; r <= bottom && refs.length < cap; r++) refs.push(encodeAddress(c, r));
    }
  }
  return refs;
}

// Rewrite named parts of a written package and read the result back — the way to feed the
// reader the hand-edited OOXML forms real producers emit but the writer itself never generates
// (an explicit-false boolean flag `<b val="0"/>`, an alignment element carrying only `wrapText="0"`,
// an injected xf). `edits` maps a part path to a (xml) => xml transform; unlisted parts pass through.
function reloadPatched(buffer, edits) {
  const files = unzipSync(buffer);
  for (const [name, transform] of Object.entries(edits)) {
    files[name] = strToU8(transform(strFromU8(files[name])));
  }
  return readXlsx(zipSync(files));
}

// Parse an XML tag's attributes into a plain { name: value } map. base64 salt/hash values use
// only XML-safe characters, so a naive quoted-value scan is sufficient here.
function attrsOf(tag) {
  const out = {};
  const re = /([\w:]+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(tag)) !== null) out[m[1]] = m[2];
  return out;
}

const impl = {
  name: 'rewrite',

  decodeAddress(reference) {
    return decodeAddress(reference);
  },

  decodeRange(reference) {
    return decodeRange(reference);
  },

  cellColRowTypes(ref = 'B3') {
    const sheet = new Workbook().addWorksheet('S');
    const cell = sheet.getCell(ref);
    cell.value = 'x';
    return {col: cell.col, row: cell.row, colType: typeof cell.col, rowType: typeof cell.row};
  },

  inspectPackage(spec) {
    return packageFacts(spec, partMapOf(writeXlsx(buildFrom(spec))));
  },

  // Author a pivot table over source data containing XML-special characters (& < > " ') and a
  // missing field value, write, and report whether the emitted pivotCacheDefinition is well-formed
  // and free of raw unescaped ampersands. Mirrors the oracle's shape → { ok, writeError,
  // cacheWellFormed, hasRawUnescapedAmp }.
  pivotCacheSpecialCharsReport() {
    try {
      const wb = new Workbook();
      const src = wb.addWorksheet('Data');
      src.getCell('A1').value = 'Name';
      src.getCell('B1').value = 'Region';
      src.getCell('C1').value = 'Amount';
      src.getCell('A2').value = 'Smith & Co';
      src.getCell('B2').value = '<West>';
      src.getCell('C2').value = 10;
      src.getCell('B3').value = 'East';
      src.getCell('C3').value = 20;
      src.getCell('A4').value = 'It\'s "best"';
      src.getCell('B4').value = 'West';
      src.getCell('C4').value = 30;
      wb.addWorksheet('Pivot').addPivotTable({
        source: src,
        rows: ['Name'],
        columns: ['Region'],
        values: ['Amount'],
        metric: 'sum',
      });
      const parts = partMapOf(writeXlsx(wb));
      const key = Object.keys(parts).find(n => /pivotCacheDefinition\d*\.xml$/.test(n));
      const cacheXml = key ? parts[key] : '';
      const hasRawUnescapedAmp = /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/.test(cacheXml);
      return {ok: true, writeError: null, cacheWellFormed: cacheXml ? !hasRawUnescapedAmp : false, hasRawUnescapedAmp};
    } catch (e) {
      return {ok: false, writeError: String((e && e.message) || e), cacheWellFormed: null, hasRawUnescapedAmp: null};
    }
  },

  // Set an autofilter over a range, write, and report the sheet's autoFilter ref plus whether the
  // workbook declares the hidden, sheet-scoped `_xlnm._FilterDatabase` defined name portable consumers
  // (LibreOffice) rely on → { autoFilterRef, hasFilterDatabase, filterDatabaseHidden,
  // filterDatabaseFormula }. Mirrors the oracle's shape.
  autoFilterDefinedNameReport(ref = 'A1:B2') {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'H1';
    sheet.getCell('B1').value = 'H2';
    sheet.getCell('A2').value = 1;
    sheet.getCell('B2').value = 2;
    sheet.autoFilter = ref;
    const parts = partMapOf(writeXlsx(workbook));
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const wbXml = parts['xl/workbook.xml'] || '';
    const autoFilterRef = (sheetXml.match(/<autoFilter\b[^>]*ref="([^"]*)"/) || [])[1] ?? null;
    const filterDb = wbXml.match(/<definedName\b([^>]*)name="_xlnm._FilterDatabase"([^>]*)>([\s\S]*?)<\/definedName>/);
    return {
      autoFilterRef,
      hasFilterDatabase: !!filterDb,
      filterDatabaseHidden: filterDb ? /hidden="1"/.test(filterDb[1] + filterDb[2]) : false,
      filterDatabaseFormula: filterDb ? filterDb[3] : null,
    };
  },

  // Drive the streaming writer to produce a package, then treat the bytes as an UNTRUSTED archive:
  // JSZip (an independent zip impl) extracts with CRC checking on, so a mismatched entry throws. Also
  // report part count, empty parts, and a whole-file re-read → { partCount, emptyParts, crcValid,
  // crcError, reloadOk, reloadError, sheetNames, firstCol }. The streamed container must be valid with
  // no repair step.
  async streamWritePackageReport({rows = 50} = {}) {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    for (let i = 1; i <= rows; i++) sheet.addRow([`r${i}`, i]).commit();
    sheet.commit();
    const buffer = Buffer.from(await writer.commit());

    let crcValid = true;
    let crcError = null;
    const emptyParts = [];
    let partCount = 0;
    try {
      const zip = await JSZip.loadAsync(buffer, {checkCRC32: true});
      const names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
      partCount = names.length;
      for (const n of names) {
        const bytes = await zip.files[n].async('nodebuffer');
        if (bytes.length === 0) emptyParts.push(n);
      }
    } catch (e) {
      crcValid = false;
      crcError = String((e && e.message) || e);
    }

    let reloadOk = true;
    let reloadError = null;
    let sheetNames = [];
    let firstCol = [];
    try {
      const wb = readXlsx(buffer);
      sheetNames = wb.worksheets.map(s => s.name);
      const s = wb.worksheets[0];
      for (let i = 1; i <= Math.min(rows, 3); i++) firstCol.push(normalizeStreamValue(s.getCell(`A${i}`).value));
    } catch (e) {
      reloadOk = false;
      reloadError = String((e && e.message) || e);
    }
    return {partCount, emptyParts, crcValid, crcError, reloadOk, reloadError, sheetNames, firstCol};
  },

  // Drive the streaming worksheet writer through row ops (addRow/addRows), commit, and read the
  // requested cells back → { ok, error, cells, rowCount }. Exercises the batch-add convenience and the
  // single-row control on the same path.
  async streamWriteSheet({ops = [], read = [], useSharedStrings = false} = {}) {
    const toRow = values => (values || []).map(specValueToModel);
    const writer = new WorkbookStreamWriter({useSharedStrings});
    let error = null;
    try {
      const sheet = writer.addWorksheet('S');
      for (const op of ops) {
        if (op.op === 'addRow') sheet.addRow(toRow(op.value)).commit();
        else if (op.op === 'addRows') sheet.addRows((op.value || []).map(toRow));
        else throw new Error(`unknown stream op: ${op.op}`);
      }
      sheet.commit();
    } catch (e) {
      error = String((e && e.message) || e);
    }
    const buffer = Buffer.from(await writer.commit());
    if (error) return {ok: false, error, cells: {}, rowCount: 0};

    const s = readXlsx(buffer).worksheets[0];
    const cells = {};
    for (const ref of read) cells[ref] = modelValueToSpec(normalizeStreamValue(s.getCell(ref).value));
    return {ok: true, error: null, cells, rowCount: s.rowCount};
  },

  // Write one string cell with the buffered writer's useSharedStrings option and report how it was
  // stored → { hasSharedStringsPart, isSharedRef, isInline }. The option must actually control
  // storage: enabled emits a sharedStrings part and a t="s" cell reference; disabled keeps the string
  // inline with no such part.
  sharedStringsOption(useSharedStrings) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'shared-me';
    const parts = partMapOf(writeXlsx(wb, {useSharedStrings}));
    const sheet = parts['xl/worksheets/sheet1.xml'] || '';
    return {
      hasSharedStringsPart: 'xl/sharedStrings.xml' in parts,
      isSharedRef: /t="s"><v>\d+<\/v>/.test(sheet),
      isInline: /t="inlineStr"><is><t>shared-me<\/t>/.test(sheet),
    };
  },

  // Commit a streaming worksheet, then add a row → { rejected, error, legibleRejection, internalCrash,
  // reloadOk }. A post-commit add must be rejected with a legible "already committed" error, not an
  // internal null-property crash, and a cleanly-committed workbook still reads back.
  async streamAddRowAfterCommit() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.addRow(['a']).commit();
    sheet.commit();
    let error = null;
    try {
      sheet.addRow(['b']).commit();
    } catch (e) {
      error = String((e && e.message) || e);
    }
    const buffer = Buffer.from(await writer.commit());
    const legibleRejection = error != null && /commit|committed|finaliz|closed/i.test(error);
    const internalCrash = error != null && /Cannot read propert|of (null|undefined)/i.test(error);
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {rejected: error != null, error, legibleRejection, internalCrash, reloadOk};
  },

  // Stream-write a sheet carrying both an autofilter and sheet protection, then inspect the emitted
  // worksheet XML for CT_Worksheet ordering → { protectThrew, sheetProtectionBeforeAutoFilter,
  // reloadOk }. <sheetProtection> must precede <autoFilter>; the shared serializer guarantees it.
  async streamAutoFilterProtectionOrder() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.addRow(['H1', 'H2']).commit();
    sheet.addRow(['a', 'b']).commit();
    sheet.autoFilter = 'A1:B1';
    let protectThrew = false;
    try {
      sheet.protect('pw', {});
    } catch {
      protectThrew = true;
    }
    sheet.commit();
    const buffer = Buffer.from(await writer.commit());
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const posProt = xml.indexOf('<sheetProtection');
    const posAf = xml.indexOf('<autoFilter');
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {protectThrew, sheetProtectionBeforeAutoFilter: posProt >= 0 && posAf >= 0 && posProt < posAf, reloadOk};
  },

  // Probe the streaming writer's image parity → { writerAddImage, sheetAddImage, error, mediaParts,
  // drawingParts }. A registered image anchored on a streamed sheet must embed a media binary and a
  // drawing part, exactly like the in-memory writer. The oracle anchors by range string; the model's
  // addImage takes grid points, so decode the range into a tl/br rectangle here.
  async streamWriterImageSupport(range = 'B2:D6') {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    const surface = {
      writerAddImage: typeof writer.addImage === 'function',
      sheetAddImage: typeof sheet.addImage === 'function',
    };
    let error = null;
    let buffer = null;
    try {
      const imageId = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
      const {left, top, right, bottom} = decodeRange(range);
      sheet.addImage(imageId, {tl: {col: left - 1, row: top - 1}, br: {col: right, row: bottom}});
      sheet.addRow(['x']).commit();
      sheet.commit();
      buffer = Buffer.from(await writer.commit());
    } catch (e) {
      error = String((e && e.message) || e);
    }
    let mediaParts = [];
    let drawingParts = [];
    if (!error && buffer) {
      const parts = Object.keys(partMapOf(buffer));
      mediaParts = parts.filter(n => /xl\/media\//.test(n));
      drawingParts = parts.filter(n => /drawing/.test(n));
    }
    return {...surface, error, mediaParts, drawingParts};
  },

  // Build a workbook whose sheets place images at the spec's ranges, write it, and report the
  // serialized drawing-anchor geometry (type, editAs, from/to, one-cell extent, spPr transform) as
  // plain numbers — the surface a case asserts against for anchor correctness.
  inspectImageAnchors(spec) {
    const parts = partMapOf(writeXlsx(buildFrom(spec)));
    const drawingParts = Object.keys(parts).filter(f => /^xl\/drawings\/drawing\d+\.xml$/.test(f)).sort();
    const anchors = [];
    let xmlOk = true;
    for (const p of drawingParts) {
      const xml = parts[p];
      if (!imageXmlWellFormed(xml)) xmlOk = false;
      for (const m of xml.matchAll(/<xdr:(oneCellAnchor|twoCellAnchor)\b([^>]*)>([\s\S]*?)<\/xdr:\1>/g)) {
        const body = m[3];
        const fromBlock = (body.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/) || [])[1];
        const toBlock = (body.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/) || [])[1];
        const extTag = body.match(/<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
        const editAs = (m[2].match(/editAs="([^"]*)"/) || [])[1] || null;
        const sppr = (body.match(/<xdr:spPr\b[\s\S]*?<\/xdr:spPr>/) || [])[0] || '';
        const offTag = sppr.match(/<a:off\b[^>]*x="(-?\d+)"[^>]*y="(-?\d+)"/);
        const spExtTag = sppr.match(/<a:ext\b[^>]*cx="(-?\d+)"[^>]*cy="(-?\d+)"/);
        const off = offTag ? {x: Number(offTag[1]), y: Number(offTag[2])} : null;
        const spExt = spExtTag ? {cx: Number(spExtTag[1]), cy: Number(spExtTag[2])} : null;
        anchors.push({
          anchorType: m[1] === 'oneCellAnchor' ? 'oneCell' : 'twoCell',
          editAs,
          from: parseAnchorSide(fromBlock),
          to: parseAnchorSide(toBlock),
          ext: extTag ? {cx: Number(extTag[1]), cy: Number(extTag[2])} : null,
          spPr: {
            hasXfrm: /<a:xfrm\b/.test(sppr),
            off,
            ext: spExt,
            zeroedTransform: !!(off && spExt && off.x === 0 && off.y === 0 && spExt.cx === 0 && spExt.cy === 0),
          },
        });
      }
    }
    return {anchors, drawingCount: drawingParts.length, xmlWellFormed: xmlOk};
  },

  // Anchor two images to single-cell ranges (C2, C3), interleaving cell writes, and report each
  // anchor's serialized from col/row → { anchorCount, froms }. A cell-range anchor must resolve to
  // that exact cell with no off-by-one row drift.
  cellAnchoredImagePositionReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'r1';
    anchorSpecImage(ws, wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), 'C2:C2');
    ws.getCell('A2').value = 'r3';
    anchorSpecImage(ws, wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), 'C3:C3');
    const drawingXml = partMapOf(writeXlsx(wb))['xl/drawings/drawing1.xml'] || '';
    const froms = [...drawingXml.matchAll(/<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/g)].map(m => ({
      col: Number(m[1]),
      row: Number(m[2]),
    }));
    return {anchorCount: froms.length, froms};
  },

  // Anchor a two-cell and a one-cell image, round-trip, and report each read-back image's top-left
  // cell and whether its media survived → { count, images, mediaCount }.
  enumerateImagesAfterRoundtrip() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.addImage(wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), {tl: {col: 1, row: 1}, br: {col: 3, row: 3}});
    ws.addImage(wb.addImage({buffer: ONE_PX_PNG, extension: 'png'}), {tl: {col: 5, row: 5}, ext: {width: 50, height: 50}});
    const reread = readXlsx(writeXlsx(wb));
    const images = (reread.getWorksheet('S')?.images || []).map(im => {
      const from = im.anchor.from;
      return {tl: from ? {col: from.col, row: from.row} : null, hasMedia: !!reread.getImage(im.imageId)};
    });
    return {count: images.length, images, mediaCount: reread.media.length};
  },

  // Register two DISTINCT images (told apart by byte length) and place them interleaved (default
  // B, A, A). Resolve, per anchor, which media part its embed actually references → so a case can
  // assert every anchor renders the image it was placed with, including a reused image.
  interleavedImageAnchors(placement = 'BAA') {
    const PNG_A = hexBytes(
      '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da6360000002000001e5273db40000000049454e44ae426082'
    );
    const PNG_B = hexBytes(
      '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154789c6300010000050001'
    );
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const ids = {A: wb.addImage({buffer: PNG_A, extension: 'png'}), B: wb.addImage({buffer: PNG_B, extension: 'png'})};
    const placed = [...placement];
    placed.forEach((letter, i) => {
      const col = i * 2;
      sheet.addImage(ids[letter], {tl: {col, row: 0}, br: {col: col + 2, row: 2}});
    });
    const buffer = writeXlsx(wb);
    const raw = unzipSync(buffer);
    const relsXml = strFromU8(raw['xl/drawings/_rels/drawing1.xml.rels'] || new Uint8Array());
    const relTarget = {};
    for (const t of relsXml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
      const a = attrsOf(t[0]);
      relTarget[a.Id] = (a.Target || '').split('/').pop();
    }
    const drawingXml = strFromU8(raw['xl/drawings/drawing1.xml'] || new Uint8Array());
    const embedOrder = [...drawingXml.matchAll(/r:embed="([^"]*)"/g)].map(m => m[1]);
    const resolvedMedia = embedOrder.map(rid => relTarget[rid] ?? null);
    const mediaSizes = {};
    for (const name of Object.keys(raw)) {
      const m = name.match(/^xl\/media\/(image\d+\.png)$/);
      if (m) mediaSizes[m[1]] = raw[name].length;
    }
    const resolvedLetter = resolvedMedia.map(media => {
      const size = mediaSizes[media];
      if (size === PNG_A.length) return 'A';
      if (size === PNG_B.length) return 'B';
      return '?';
    });
    return {
      placed,
      embedOrder,
      resolvedMedia,
      resolvedLetter,
      distinctMediaCount: Object.keys(mediaSizes).length,
      distinctRelTargets: new Set(Object.values(relTarget)).size,
    };
  },

  // Load a workbook from bytes, register and anchor an image on the loaded worksheet, re-serialize,
  // and report the media/drawing presence and re-read image count → locks that adding an image to a
  // *loaded* (not freshly created) worksheet persists.
  addImageToLoadedWorksheetReport(range = 'B2:C4') {
    const base = new Workbook();
    base.addWorksheet('S').getCell('A1').value = 'x';
    const loaded = readXlsx(writeXlsx(base));
    anchorSpecImage(loaded.getWorksheet('S'), loaded.addImage({buffer: ONE_PX_PNG, extension: 'png'}), range);
    const out = writeXlsx(loaded);
    const files = Object.keys(partMapOf(out));
    const reloadImages = readXlsx(out).getWorksheet('S')?.images || [];
    return {
      hasMedia: files.some(f => /xl\/media\//.test(f)),
      hasDrawing: files.some(f => /xl\/drawings\/drawing/.test(f)),
      reloadImageCount: reloadImages.length,
    };
  },

  // Read a fixture and report each image's normalized anchor range → for asserting a file whose
  // drawing anchors were authored as cell ranges reads without throwing and exposes integer cell
  // coordinates, never a raw string.
  readFixtureImageAnchors(rel) {
    const workbook = readFixture(rel);
    const images = [];
    for (const sheet of workbook.worksheets) {
      for (const im of sheet.images) {
        const from = im.anchor.from;
        const to = im.anchor.to;
        images.push({
          sheet: sheet.name,
          editAs: im.anchor.editAs ?? null,
          tl: from ? {col: from.col, row: from.row} : null,
          br: to ? {col: to.col, row: to.row} : null,
        });
      }
    }
    return {images, count: images.length};
  },

  // Add one image whose extension may carry a leading dot or a query string, write, and report the
  // media filenames and re-read image count → a dirty extension must sanitise to a well-formed media
  // name the reader still recognises, not a doubled-separator name that drops the image.
  imageExtensionRoundtrip(extension = 'png') {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const id = wb.addImage({buffer: ONE_PX_PNG, extension});
    ws.addImage(id, {tl: {col: 1, row: 1}, br: {col: 3, row: 3}});
    const buffer = writeXlsx(wb);
    const mediaParts = Object.keys(partMapOf(buffer))
      .filter(n => /^xl\/media\/.+/.test(n))
      .map(n => n.replace(/^xl\/media\//, ''));
    const images = readXlsx(buffer).getWorksheet('S')?.images || [];
    return {
      mediaParts,
      doubledSeparator: mediaParts.some(n => /\.\./.test(n)),
      reloadedImageCount: images.length,
    };
  },

  // Anchor two images, then remove one by its media id → { supported, before, after, removedGone,
  // othersSurvive }. Removal must drop exactly the targeted image and leave the rest anchored.
  removeImageReport(range = 'A1:B2') {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    const id1 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    const id2 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    anchorSpecImage(ws, id1, range);
    anchorSpecImage(ws, id2, 'C1:D2');
    const supported = typeof ws.removeImage === 'function';
    const before = ws.images.length;
    let after = before;
    let removedGone = false;
    let othersSurvive = false;
    if (supported) {
      ws.removeImage(id1);
      const ids = ws.images.map(i => i.imageId);
      after = ws.images.length;
      removedGone = !ids.includes(id1);
      othersSurvive = ids.includes(id2);
    }
    return {supported, before, after, removedGone, othersSurvive};
  },

  // Anchor an image and append rows in both orders → { imageFirst, rowsFirst }, each { rowCount,
  // firstDataCell }. Anchoring a floating image is metadata overlay, not row insertion: it must not
  // advance the row-append cursor, so the layout is identical regardless of add order.
  imageAnchorRowAppendReport() {
    const run = order => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
      if (order === 'image-first') {
        anchorSpecImage(sheet, id, 'A1:B3');
        sheet.addRows([['a'], ['b'], ['c']]);
      } else {
        sheet.addRows([['a'], ['b'], ['c']]);
        anchorSpecImage(sheet, id, 'A1:B3');
      }
      return {rowCount: sheet.rowCount, firstDataCell: sheet.getCell('A1').value ?? null};
    };
    return {imageFirst: run('image-first'), rowsFirst: run('rows-first')};
  },

  // Read a fixture, load-and-rewrite it, and report the picture's drawing-anchor rotation before and
  // after → { sourceRot, rewrittenRot }. An image rotation (rot on <a:xfrm>) must survive the round-trip.
  roundtripFixtureImageRotation(rel) {
    const rotOf = xml => {
      const m = xml.match(/<a:xfrm\b[^>]*\brot="(-?\d+)"/);
      return m ? Number(m[1]) : null;
    };
    const drawingName = parts => Object.keys(parts).find(f => /^xl\/drawings\/drawing\d+\.xml$/.test(f));
    const srcParts = partMapOf(fixtureBytes(rel));
    const srcDrawing = drawingName(srcParts);
    const sourceRot = srcDrawing ? rotOf(srcParts[srcDrawing]) : null;
    const outParts = partMapOf(writeXlsx(readXlsx(fixtureBytes(rel))));
    const outDrawing = drawingName(outParts);
    const rewrittenRot = outDrawing ? rotOf(outParts[outDrawing]) : null;
    return {sourceRot, rewrittenRot};
  },

  // Read a fixture, write it back unchanged, and report package-part facts before/after →
  // { source, rewritten } — for asserting a no-op round-trip PRESERVES parts the reader does not
  // model (a vector-shape drawing, a header/footer image and its VML) instead of dropping them.
  roundtripFixturePackageParts(rel) {
    const source = packagePartFacts(partMapOf(fixtureBytes(rel)));
    const rewritten = packagePartFacts(partMapOf(writeXlsx(readXlsx(fixtureBytes(rel)))));
    return {source, rewritten};
  },

  // Author a sheet, round-trip, load, append more rows after the last populated row, round-trip again →
  // { loadedRowCount, finalRowCount, rows }. The load-bearing fact: a reloaded sheet reports its last
  // populated row so addRow lands new content at N+1 with no gap or overwrite.
  appendRowsAfterReload(initial = [], append = []) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const row of initial) sheet.addRow(row);

    const loaded = readXlsx(writeXlsx(workbook));
    const s = loaded.getWorksheet('S');
    const loadedRowCount = s.rowCount;
    for (const row of append) s.addRow(row);

    const final = readXlsx(writeXlsx(loaded));
    const f = final.getWorksheet('S');
    // Mirror the oracle's `row.values.slice(1)` per-row array: each row is sized to its own populated
    // extent, holes are null, and an empty row is an empty array — indexed by row number so a gap shows.
    const rows = Array.from({length: f.rowCount}, () => []);
    for (const {number, cells} of f.rows()) {
      const maxCol = cells.reduce((m, c) => Math.max(m, c.col), 0);
      const arr = new Array(maxCol).fill(null);
      for (const cell of cells) arr[cell.col - 1] = normalizeStreamValue(cell.value);
      rows[number - 1] = arr;
    }
    return {loadedRowCount, finalRowCount: f.rowCount, rows};
  },

  // Read a fixture's single _xlnm.Print_Area name (a comma-separated range list), re-write it, and read
  // it again → { sourceRangeCount, readPrintArea, rewrittenRangeCount }. Both disjoint ranges must
  // survive read and re-serialization, never truncated to the first.
  roundtripFixturePrintAreas(rel) {
    const printAreaOf = wb => wb.definedNames.find(n => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const source = readXlsx(fixtureBytes(rel));
    const readPrintArea = printAreaOf(source);
    const sourceRangeCount = readPrintArea.split(',').filter(Boolean).length;
    const rewritten = readXlsx(writeXlsx(source));
    const rewrittenRangeCount = printAreaOf(rewritten).split(',').filter(Boolean).length;
    return {sourceRangeCount, readPrintArea, rewrittenRangeCount};
  },

  // Author a sheet-scoped _xlnm.Print_Area over a comma-separated area, round-trip, and report the
  // emitted ranges (sheet prefix stripped) → { ranges }. Two disjoint areas must emit two proper
  // rectangular ranges in one name, not a truncated single range.
  writePrintAreaDefinedName(area) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    workbook.defineName({name: '_xlnm.Print_Area', scope: sheet.name, refersTo: printAreaRefersTo(sheet.name, area)});
    const back = readXlsx(writeXlsx(workbook));
    const refersTo = back.definedNames.find(n => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const ranges = refersTo.split(',').map(r => r.split('!').pop());
    return {ranges};
  },

  // Author a sheet-scoped _xlnm.Print_Area over one area (whole-column or bounded), round-trip, and
  // report the written and recovered forms → { writtenDefinedName, reReadPrintArea, reloadOk }. A
  // column-only reference ($A:$D) must recover intact, never decoded to a NaN-mangled address.
  printAreaRoundtrip(area) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    workbook.defineName({name: '_xlnm.Print_Area', scope: sheet.name, refersTo: printAreaRefersTo(sheet.name, area)});
    let reloadOk = true;
    let back;
    try {
      back = readXlsx(writeXlsx(workbook));
    } catch {
      reloadOk = false;
    }
    const refersTo = back?.definedNames.find(n => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const reReadPrintArea = refersTo.split('!').pop()?.replace(/\$/g, '') ?? '';
    return {writtenDefinedName: refersTo, reReadPrintArea, reloadOk};
  },

  // Freeze the first row, write, and report the emitted pane plus a round-trip → { paneEmitted,
  // reReadState, reReadYSplit, reReadXSplit }. A frozen-header view serializes a <pane> and reloads
  // as a frozen split of one row and no columns.
  frozenTopRowRoundtrip() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const paneEmitted = /<pane\b[^>]*ySplit="1"[^>]*state="frozen"/.test(sheetXml);
    const view = readXlsx(buffer).getWorksheet('S').view;
    return {
      paneEmitted,
      reReadState: view.state ?? 'normal',
      reReadYSplit: view.ySplit ?? 0,
      reReadXSplit: view.xSplit ?? 0,
    };
  },

  // Freeze a view, then unfreeze it, and report the pane presence in each written form plus a reload
  // → { frozenHasPane, normalHasPane, reloadedState, reloadedHasSplit }. Unfreezing must leave no
  // leftover <pane> (which triggers Excel's repair prompt) and reload as a normal, unsplit view.
  unfreezeViewRoundtrip() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1);
    const frozenHasPane = /<pane\b/.test(partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '');

    sheet.unfreeze();
    const normalBuffer = writeXlsx(workbook);
    const normalHasPane = /<pane\b/.test(partMapOf(normalBuffer)['xl/worksheets/sheet1.xml'] || '');
    const view = readXlsx(normalBuffer).getWorksheet('S').view;
    return {
      frozenHasPane,
      normalHasPane,
      reloadedState: view.state ?? 'normal',
      reloadedHasSplit: (view.xSplit ?? 0) > 0 || (view.ySplit ?? 0) > 0,
    };
  },

  // Author three columns with distinct widths (one hidden), write, then REVERSE the order of the
  // emitted `<col>` tags — the shape foreign generators (excelize, jxls-poi) produce — and read the
  // patched package back → { w1, w2, w3, hidden2 }. Each column's width and hidden flag must bind to
  // the column its min/max names, regardless of document order.
  outOfOrderColumnsReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getColumn(1).width = 10;
    sheet.getColumn(2).width = 20;
    sheet.getColumn(2).hidden = true;
    sheet.getColumn(3).width = 30;
    const back = reloadPatched(writeXlsx(wb), {
      'xl/worksheets/sheet1.xml': xml =>
        xml.replace(/<cols>([\s\S]*?)<\/cols>/, (_, inner) => {
          const tags = inner.match(/<col\b[^>]*\/>/g) || [];
          return `<cols>${tags.reverse().join('')}</cols>`;
        }),
    }).getWorksheet('S');
    return {
      w1: back.getColumn(1).width,
      w2: back.getColumn(2).width,
      w3: back.getColumn(3).width,
      hidden2: back.getColumn(2).hidden ?? false,
    };
  },

  // Inject a `<f t="dataTable">` into a written sheet, read it back, and re-write → { reloadOk,
  // readShareType, readRef, readResult, outHasDataTable }. The reader must surface the data-table
  // kind/range/result, and a read-modify-write must re-emit t="dataTable" rather than dropping it.
  dataTableFormulaRoundtrip() {
    const seed = new Workbook();
    const seedSheet = seed.addWorksheet('S');
    seedSheet.getCell('A1').value = 1;
    seedSheet.getCell('B1').value = 2;
    seedSheet.getCell('B2').value = 99;
    const parts = unzipSync(writeXlsx(seed));
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
        /<c r="B2"[^>]*>[\s\S]*?<\/c>/,
        '<c r="B2"><f t="dataTable" ref="B2:B5" dt2D="0" dtr="1" r1="A1"/><v>99</v></c>'
      )
    );
    const injected = zipSync(parts);

    let reloadOk = false;
    let readShareType = null;
    let readRef = null;
    let readResult = null;
    let outHasDataTable = false;
    try {
      const reload = readXlsx(injected);
      const value = reload.getWorksheet('S').getCell('B2').value;
      if (value && typeof value === 'object') {
        readShareType = value.shareType ?? null;
        readRef = value.ref ?? null;
        readResult = value.result ?? null;
      }
      reloadOk = true;
      const outXml = strFromU8(unzipSync(writeXlsx(reload))['xl/worksheets/sheet1.xml']);
      outHasDataTable = /t="dataTable"/.test(outXml);
    } catch {
      reloadOk = false;
    }
    return {readShareType, readRef, readResult, reloadOk, outHasDataTable};
  },

  // Author a date-type validation whose operand coerces to a serial (or fails to), write, and report
  // the emitted first bound → { formula1, hasNaN }. A real date writes a numeric serial; a
  // non-coercible operand must drop the bound, never serialize the literal "NaN".
  authorDateValidation(operand) {
    const serial = (() => {
      const ms = Date.parse(operand);
      return Number.isNaN(ms) ? Number.NaN : ms / 86_400_000 + 25_569; // Unix epoch → 1900 serial
    })();
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addDataValidation('A1', {type: 'date', operator: 'greaterThan', formulae: [serial]});
    const sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    const formula1 = (sheetXml.match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
    return {formula1, hasNaN: /NaN/.test(sheetXml)};
  },

  // Apply one list validation with a cross-sheet source over a vertical span, round-trip, and report
  // the per-row source references plus how many dataValidation blocks were emitted → { source,
  // formulae, allIdentical, sqrefBlocks }. Every row must keep the exact source (no relative drift),
  // and the identical rules collapse into one spanning sqref.
  listValidationSourceRangeAcrossRows(rows, source) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addDataValidation(`A1:A${rows}`, {type: 'list', formulae: [source]}, {extended: true});
    const buffer = writeXlsx(workbook);
    const reloaded = readXlsx(buffer).getWorksheet('S');
    const formulae = [];
    for (let r = 1; r <= rows; r++) formulae.push(reloaded.dataValidationAt(`A${r}`)?.formulae?.[0] ?? null);
    const allIdentical = formulae.every(f => f === source);
    const dvXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const sqrefBlocks = (dvXml.match(/<(?:x14:)?dataValidation[\s>]/g) || []).length;
    return {source, formulae, allIdentical, sqrefBlocks};
  },

  // Write a non-finite numeric cell (NaN / Infinity / -Infinity) and report whether the sheet XML
  // carries a bare token in a <v> → { hasNonFiniteToken, token }. A non-finite value has no OOXML
  // representation, so it must serialize as a valueless cell, never a literal "NaN"/"Infinity".
  nonFiniteCellReport(kind) {
    const value = kind === 'NaN' ? Number.NaN : kind === '-Infinity' ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY;
    const workbook = new Workbook();
    workbook.addWorksheet('S').getCell('A1').value = value;
    const sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    const token = (sheetXml.match(/<c r="A1"[^>]*>\s*<v>([\s\S]*?)<\/v>/) || [])[1] ?? null;
    return {hasNonFiniteToken: /<v>[^<]*(NaN|Infinity)[^<]*<\/v>/.test(sheetXml), token};
  },

  // Write a value under a date number format and report the sheet XML's health → { ok, hasNaN,
  // hasInvalidDate, cellXml }. A string, a null (empty cell), or an Invalid Date under a date format
  // must never leak a bare "NaN" or "Invalid Date" into the cell value.
  dateNumFmtValueReport(kind) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell('A1');
    if (kind === 'string') cell.value = 'not a date';
    else if (kind === 'null') cell.value = null;
    else if (kind === 'invalidDate') cell.value = new Date(NaN);
    cell.numFmt = 'yyyy-mm-dd';
    let ok = true;
    let sheetXml = '';
    try {
      sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    } catch {
      ok = false;
    }
    const cellXml = (sheetXml.match(/<c r="A1"[\s\S]*?(?:\/>|<\/c>)/) || [])[0] ?? '';
    return {ok, hasNaN: /NaN/.test(sheetXml), hasInvalidDate: /Invalid Date/.test(sheetXml), cellXml};
  },

  // Read a fixture, write it back, and parse the requested cells straight from the re-emitted sheet
  // XML → { hasNaNToken, cells }. Each cell is { t, formula, value } read off the raw `<c>`. Guards
  // that a string-typed formula result under a date format is not coerced to a numeric/NaN cell.
  roundtripFixtureCellXml(rel, refs = []) {
    const parts = partMapOf(writeXlsx(readXlsx(fixtureBytes(rel))));
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const cells = {};
    for (const ref of refs) {
      const match = sheetXml.match(new RegExp(`<c r="${ref}"([^>]*)>([\\s\\S]*?)</c>`));
      if (!match) continue;
      const t = (match[1].match(/\bt="([^"]*)"/) || [])[1] ?? null;
      const formula = (match[2].match(/<f[^>]*>([\s\S]*?)<\/f>/) || [])[1] ?? null;
      const rawValue = (match[2].match(/<v>([\s\S]*?)<\/v>/) || [])[1] ?? null;
      // A t="str" (or shared-string) cell holds text; anything else with a bare <v> is numeric.
      const value = rawValue === null ? null : t === 'str' ? rawValue : Number(rawValue);
      cells[ref] = {t, formula, value};
    }
    return {hasNaNToken: /<v>[^<]*NaN[^<]*<\/v>/.test(sheetXml), cells};
  },

  // Round-trip formula cells whose cached results are truthy and falsy (2, 0, false, '') and report
  // each recovered result → { truthy, zero, boolFalse, emptyString } of { hasResult, result }. A falsy
  // result (0, false, empty string) must survive, not be dropped as if the formula had no cached value.
  formulaFalsyResultReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = {formula: '1+1', result: 2};
    sheet.getCell('A2').value = {formula: 'B1-B1', result: 0};
    sheet.getCell('A3').value = {formula: 'FALSE()', result: false};
    sheet.getCell('A4').value = {formula: 'T("")', result: ''};
    const back = readXlsx(writeXlsx(workbook)).getWorksheet('S');
    const probe = ref => {
      const value = back.getCell(ref).value;
      const hasResult = !!value && typeof value === 'object' && 'result' in value;
      return {hasResult, result: hasResult ? value.result : undefined};
    };
    return {truthy: probe('A1'), zero: probe('A2'), boolFalse: probe('A3'), emptyString: probe('A4')};
  },

  // Round-trip a formula whose cached result is a Date → { isValidDate, resultIso, keepsFormula }. The
  // date result reads back as a valid Date (the default date format survives), and the cell stays a
  // formula cell rather than collapsing to a bare value.
  formulaDateResultReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = {formula: 'TODAY()', result: new Date(Date.UTC(2021, 0, 2))};
    const value = readXlsx(writeXlsx(workbook)).getWorksheet('S').getCell('A1').value;
    const result = value && typeof value === 'object' ? value.result : undefined;
    const isValidDate = result instanceof Date && !Number.isNaN(result.getTime());
    return {
      isValidDate,
      resultIso: result instanceof Date ? isoOrNull(result) : String(result),
      keepsFormula: !!value && typeof value === 'object' && typeof value.formula === 'string',
    };
  },

  // Add a sheet, then probe name lookup and uniqueness for case-consistency → { foundExact,
  // foundVariant, addVariantThrew }. Lookup and add must agree on identity: a case-variant name is
  // found by getWorksheet AND rejected by addWorksheet (both case-insensitive), so no absent-yet-
  // unaddable surprise exists.
  worksheetNameLookupReport() {
    const workbook = new Workbook();
    workbook.addWorksheet('Sheet');
    const foundExact = workbook.getWorksheet('Sheet') !== undefined;
    const foundVariant = workbook.getWorksheet('sheet') !== undefined;
    let addVariantThrew = false;
    try {
      workbook.addWorksheet('sheet');
    } catch {
      addVariantThrew = true;
    }
    return {foundExact, foundVariant, addVariantThrew};
  },

  // Append rows in every shape (dense array, sparse array, keyed object, mixed batch), round-trip, and
  // read them back letter-keyed by row number → { rows }. Column keys bind object values to columns;
  // dense/sparse arrays map positionally with holes left empty; a numeric/date value survives typed.
  appendRowShapes() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getColumn(1).key = 'k1';
    sheet.getColumn(2).key = 'k2';
    sheet.addRow(['header']); // row 1 — keeps the checked rows at their stated numbers
    sheet.addRow(['a', 'b', 'c']); // row 2 — dense positional array
    // eslint-disable-next-line no-sparse-arrays
    sheet.addRow(['x', , 'z']); // row 3 — sparse array, gap at column B
    sheet.addRow({k1: 'o1', k2: 'o2'}); // row 4 — keyed object
    sheet.addRow([7, new Date(Date.UTC(2021, 0, 2))]); // row 5 — number + date
    sheet.addRows([['m1', 'm2'], {k1: 'n1'}]); // rows 6, 7 — mixed batch

    const loaded = readXlsx(writeXlsx(workbook));
    const s = loaded.getWorksheet('S');
    const rows = {};
    for (const {number, cells} of s.rows()) {
      const row = {};
      for (const cell of cells) row[encodeAddress(cell.col, number).match(/^[A-Z]+/)[0]] = normalizeStreamValue(cell.value);
      rows[number] = row;
    }
    // Every checked column reads as null when the round-trip left it empty, so a gap is visible.
    for (const n of Object.keys(rows)) for (const col of ['A', 'B', 'C']) rows[n][col] ??= null;
    return {rows};
  },

  // Feed addRow an array built in another realm (a vm context): Array.isArray must recognize it so its
  // elements fill columns → { isArrayCrossRealm, a, b, c }. `instanceof Array` would miss it and place
  // nothing, walking it as a keyed object instead.
  async crossRealmArrayRow() {
    const vm = await import('node:vm');
    const arr = vm.runInNewContext('[10, 20, 30]');
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addRow(arr);
    return {
      isArrayCrossRealm: Array.isArray(arr),
      a: sheet.getCell('A1').value ?? null,
      b: sheet.getCell('B1').value ?? null,
      c: sheet.getCell('C1').value ?? null,
    };
  },

  // Stream-read a styled workbook and rebuild it through the streaming writer, copying each cell's
  // value AND resolved style onto the new sheet → { copyError, loadOk, fontBold, fontColor, numFmt,
  // hasFill }. The streaming reader surfaces each cell's style facets, so the per-cell font, fill, and
  // number format survive the streaming read→write copy and the emitted styles part loads cleanly.
  async streamingStyleCopyReport() {
    const src = new Workbook();
    const c = src.addWorksheet('S').getCell('A1');
    c.value = 'hello';
    c.font = {bold: true, color: {argb: 'FFFF0000'}};
    c.numFmt = '0.00%';
    c.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
    const srcBuffer = writeXlsx(src);

    const writer = new WorkbookStreamWriter();
    let copyError = null;
    try {
      for (const sheet of readWorkbookStream(srcBuffer)) {
        const ows = writer.addWorksheet(sheet.name);
        for (const row of sheet.rows()) {
          for (const cell of row.cells) {
            const target = ows.getCell(cell.address);
            target.value = cell.value;
            if (cell.style) {
              if (cell.style.font !== undefined) target.font = cell.style.font;
              if (cell.style.fill !== undefined) target.fill = cell.style.fill;
              if (cell.style.numFmt !== undefined) target.numFmt = cell.style.numFmt;
              if (cell.style.border !== undefined) target.border = cell.style.border;
              if (cell.style.alignment !== undefined) target.alignment = cell.style.alignment;
            }
          }
        }
        ows.commit();
      }
    } catch (e) {
      copyError = String((e && e.message) || e);
    }
    const buffer = Buffer.from(await writer.commit());
    if (copyError) return {copyError, loadOk: false, fontBold: null, fontColor: null, numFmt: null, hasFill: null};

    let loadOk = true;
    let cell = null;
    try {
      cell = readXlsx(buffer).getWorksheet('S').getCell('A1');
    } catch {
      loadOk = false;
    }
    return {
      copyError,
      loadOk,
      fontBold: cell ? !!(cell.font && cell.font.bold) : null,
      fontColor: cell && cell.font && cell.font.color ? cell.font.color.argb ?? null : null,
      numFmt: cell ? cell.numFmt ?? null : null,
      hasFill: cell ? !!(cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) : null,
    };
  },

  // Pipe the streaming writer's output stream to a sink and report { pipeReturnsDestination, bytes,
  // valid }. Node's Readable.pipe(dest) must RETURN dest so `.pipe(out).on('finish', …)` composes,
  // while the piped payload still reconstitutes a valid workbook.
  async streamWriterPipeContract() {
    const writer = new WorkbookStreamWriter();
    const source = writer.stream;
    const sink = new PassThrough();
    const chunks = [];
    sink.on('data', c => chunks.push(c));
    const pipeReturn = source.pipe(sink);
    const pipeReturnsDestination = pipeReturn === sink;
    const ws = writer.addWorksheet('S');
    ws.addRow(['a', 'b']).commit();
    ws.commit();
    await writer.commit();
    await new Promise(res => setTimeout(res, 20));
    const buffer = Buffer.concat(chunks);
    let valid = false;
    try {
      valid = readXlsx(buffer).worksheets[0].getCell('A1').value === 'a';
    } catch {
      valid = false;
    }
    return {pipeReturnsDestination, bytes: buffer.length, valid};
  },

  // Request fullCalcOnLoad on the streaming writer (via calcProperties) and report whether it reaches
  // the output, versus the in-memory writer → { streamSetThrew, streamHasFlag, streamDefaultHasFlag,
  // memoryHasFlag }. Recalc-on-load must work identically on both writers.
  async streamingFullCalcOnLoadReport() {
    const streamCalc = async setFlag => {
      const writer = new WorkbookStreamWriter();
      let threw = false;
      if (setFlag) {
        try {
          writer.calcProperties.fullCalcOnLoad = true;
        } catch {
          threw = true;
        }
      }
      const sheet = writer.addWorksheet('S');
      sheet.getCell('A1').value = 1;
      sheet.commit();
      const buffer = Buffer.from(await writer.commit());
      const wbXml = strFromU8(unzipSync(buffer)['xl/workbook.xml']);
      return {threw, hasFlag: /fullCalcOnLoad="1"/.test(wbXml)};
    };
    const set = await streamCalc(true);
    const def = await streamCalc(false);

    const wb = new Workbook();
    wb.fullCalcOnLoad = true;
    wb.addWorksheet('S').getCell('A1').value = 1;
    const memXml = strFromU8(unzipSync(writeXlsx(wb))['xl/workbook.xml']);

    return {
      streamSetThrew: set.threw,
      streamHasFlag: set.hasFlag,
      streamDefaultHasFlag: def.hasFlag,
      memoryHasFlag: /fullCalcOnLoad="1"/.test(memXml),
    };
  },

  // Build via the streaming writer with a master formula + shared-formula slaves, reload →
  // { masterHasFormula, slaveResolved, slaveValue }. Streamed slaves must not be dropped to empty.
  async streamingSharedFormulaReport(rows = 10) {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('yua');
    for (let i = 1; i <= rows; i++) sheet.getCell(`A${i}`).value = i * 10;
    sheet.getCell('B1').value = {formula: 'A1*2', result: 20};
    for (let j = 2; j <= rows; j++) sheet.getCell(`B${j}`).value = {sharedFormula: 'B1'};
    sheet.commit();
    const buffer = Buffer.from(await writer.commit());

    const rs = readXlsx(buffer).getWorksheet('yua');
    const slave = rs.getCell('B3').value;
    const slaveIsEmpty = slave == null || (typeof slave === 'object' && Object.keys(slave).length === 0);
    const master = rs.getCell('B1').value;
    return {
      masterHasFormula: !!(master && typeof master === 'object' && 'formula' in master),
      slaveResolved: !slaveIsEmpty,
      slaveValue: normalizeStreamValue(slave ?? null),
    };
  },

  // Stream-write a sheet carrying both a conditional-formatting rule and a hyperlink cell, then report
  // the relative order of the emitted <conditionalFormatting> and <hyperlinks> blocks plus reload
  // success. Both writers share one worksheet serializer, so the streamed sheet emits the blocks in
  // CT_Worksheet order (conditionalFormatting before hyperlinks) rather than the reversed order the
  // upstream streaming writer produced.
  async streamWriteCfHyperlinkOrder() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
    sheet.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'expression',
          formulae: ['MOD(ROW(),2)=0'],
          style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFEEEEEE'}}},
        },
      ],
    });
    sheet.addRow(['x']).commit();
    sheet.commit();

    const buffer = Buffer.from(await writer.commit());
    const xml = strFromU8(unzipSync(buffer)['xl/worksheets/sheet1.xml']);
    const posCf = xml.indexOf('<conditionalFormatting');
    const posHl = xml.indexOf('<hyperlinks');
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      posConditionalFormatting: posCf,
      posHyperlinks: posHl,
      conditionalFormattingBeforeHyperlinks: posCf >= 0 && posHl >= 0 ? posCf < posHl : null,
      reloadOk,
    };
  },

  // Stream-write a sheet carrying both a data validation and a hyperlink cell, then report the relative
  // order of the emitted <dataValidations> and <hyperlinks> blocks plus reload success. CT_Worksheet
  // requires dataValidations before hyperlinks; the shared serializer emits them in that order on the
  // streaming path too. Companion to streamWriteCfHyperlinkOrder.
  async streamWriteDvHyperlinkOrder() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
    sheet.addDataValidation('B1', {type: 'list', allowBlank: true, formulae: ['"x,y,z"']});
    sheet.addRow(['r']).commit();
    sheet.commit();

    const buffer = Buffer.from(await writer.commit());
    const xml = strFromU8(unzipSync(buffer)['xl/worksheets/sheet1.xml']);
    const posDv = xml.indexOf('<dataValidations');
    const posHl = xml.indexOf('<hyperlinks');
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      posDataValidations: posDv,
      posHyperlinks: posHl,
      dataValidationsBeforeHyperlinks: posDv >= 0 && posHl >= 0 ? posDv < posHl : null,
      reloadOk,
    };
  },

  // Commit a streaming workbook over a caller-supplied writable (a plain PassThrough or a Duplex) and
  // report { settled, timedOut, bytes, valid }. The commit must settle within bounded time and the
  // sink must receive a complete, re-openable package — the library owes this even when it does not own
  // the stream.
  async streamCommitReport({duplex = false, timeoutMs = 4000} = {}) {
    const chunks = [];
    const stream = duplex
      ? new Duplex({read() {}, write(c, _e, cb) { chunks.push(c); cb(); }})
      : new PassThrough();
    if (!duplex) stream.on('data', c => chunks.push(c));

    const writer = new WorkbookStreamWriter({stream});
    const sheet = writer.addWorksheet('S');
    sheet.addRow(['a', 'b']).commit();
    sheet.commit();

    let settled = 'pending';
    const commit = writer.commit().then(
      () => (settled = 'resolved'),
      e => (settled = 'rejected:' + ((e && e.message) || e))
    );
    const timedOut = await Promise.race([
      commit.then(() => false),
      new Promise(res => setTimeout(() => res(true), timeoutMs)),
    ]);

    let valid = false;
    if (settled === 'resolved') {
      try {
        const back = readXlsx(Buffer.concat(chunks));
        valid = back.worksheets.length === 1 && back.worksheets[0].getCell('A1').value === 'a';
      } catch {
        valid = false;
      }
    }
    return {settled, timedOut, bytes: Buffer.concat(chunks).length, valid};
  },

  // Commit a streaming workbook to a destination that cannot be opened for writing and report
  // { outcome, rejected, carriesIoError, error }. The write stream errors on a later tick; commit must
  // reject with that I/O error rather than hanging forever.
  async streamCommitBadDestination() {
    const badPath = `${tmpdir()}/ts-xlsx-no-such-dir-${process.pid}/${'x'.repeat(300)}/out.xlsx`;
    let outcome = 'hung';
    let error = null;
    try {
      const writer = new WorkbookStreamWriter({filename: badPath});
      const sheet = writer.addWorksheet('S');
      sheet.addRow(['a']).commit();
      sheet.commit();
      await Promise.race([
        writer
          .commit()
          .then(() => {
            outcome = 'resolved';
          })
          .catch(e => {
            outcome = 'rejected';
            error = String((e && (e.code || e.message)) || e);
          }),
        new Promise(res => setTimeout(res, 5000)),
      ]);
    } catch (e) {
      outcome = 'threw-sync';
      error = String((e && (e.code || e.message)) || e);
    }
    return {
      outcome,
      rejected: outcome === 'rejected',
      carriesIoError: error != null && /ENOENT|ENAMETOOLONG|ENOTDIR|open|write/i.test(error),
      error,
    };
  },

  // Apply a font to each named cell, then read each requested cell's font back → { <ref>:
  // font|null }. Each cell owns its own font, so a font set on one cell is observable there
  // and nowhere else — the isolation this reports. In-memory, matching the contract: the
  // <fonts>-table write/read path is exercised by the io/xlsx unit tests.
  probeCellFonts({apply = [], read = []}) {
    const sheet = new Workbook().addWorksheet('sheet');
    for (const {cell, font} of apply) sheet.getCell(cell).font = font;
    const fonts = {};
    for (const address of read) fonts[address] = sheet.getCell(address).font ?? null;
    return JSON.parse(JSON.stringify(fonts));
  },

  // Write the spec and report the shared style table's size plus the index each requested cell
  // resolved to → { cellXfCount, indices: { <ref>: index|null } }. styles.xml is a SHARED table
  // referenced by index, so identically-styled cells must collapse to one <cellXfs> entry (one
  // shared index) — dedup neither inflating to one entry per cell nor over-collapsing distinct
  // styles. A cell left at the default style carries no `s` and reports null.
  styleDedupReport(spec, cells = []) {
    const parts = partMapOf(writeXlsx(buildFrom(spec)));
    const styles = parts['xl/styles.xml'] || '';
    const xfBlock = (styles.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [''])[0];
    const cellXfCount = (xfBlock.match(/<xf\b/g) || []).length;
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const indices = {};
    for (const ref of cells) {
      const m = sheetXml.match(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*\\bs="(\\d+)"`));
      indices[ref] = m ? Number(m[1]) : null;
    }
    return {cellXfCount, indices};
  },

  // Write one plain-valued, unformatted cell, round-trip it, and report the font it reads back →
  // { hasFont, fontName, fontSize }. An unstyled cell renders in the workbook default font, so it
  // must resolve a concrete name/size rather than reporting no font at all.
  unstyledCellFontReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'hello';
    const cell = readXlsx(writeXlsx(wb)).getWorksheet('S').getCell('A1');
    const font = cell.font || null;
    return {
      hasFont: !!font,
      fontName: font ? font.name ?? null : null,
      fontSize: font ? font.size ?? null : null,
    };
  },

  // Colour a sheet's tab with an 8-digit ARGB (alpha first) alongside an uncoloured sheet, then
  // report the round-trip → { tabColorArgbWritten, reReadArgb, uncoloredHasTab }. A set tab colour
  // must survive verbatim while a sheet with none must not gain a spurious one.
  tabColorRoundtrip() {
    const wb = new Workbook();
    const colored = wb.addWorksheet('Colored');
    colored.tabColor = {argb: 'FFFF0000'};
    colored.getCell('A1').value = 'x';
    const plain = wb.addWorksheet('Plain');
    plain.getCell('A1').value = 'y';
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const written = (sheetXml.match(/<tabColor\b[^>]*rgb="([^"]*)"/) || [null, null])[1];
    const reload = readXlsx(buffer);
    const coloredTab = reload.getWorksheet('Colored').tabColor;
    return {
      tabColorArgbWritten: written,
      reReadArgb: (coloredTab && coloredTab.argb) || null,
      uncoloredHasTab: !!reload.getWorksheet('Plain').tabColor,
    };
  },

  // Write a solid fill twice — once with a clean bare ARGB, once with a CSS-habit '#'-prefixed one —
  // and report the emitted <fgColor rgb="..."> for each → { validRgb, hashRgb }. Both must serialize
  // as valid 8-hex-digit values; a '#'-prefixed input must be normalized, never passed through as a
  // malformed 9-character colour.
  fillArgbHashPrefixReport() {
    const emittedFgColor = argb => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb}};
      const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
      return (stylesXml.match(/<fgColor rgb="([^"]*)"/) || [null, null])[1];
    };
    return {validRgb: emittedFgColor('FFBFBFBF'), hashRgb: emittedFgColor('#FFBFBFBF')};
  },

  // Author a solid fill with a 6-hex RGB (no alpha) and with a malformed value, and report how the
  // writer treats each → { sixHexRgb, rejectsMalformed }. A 6-hex RGB is the common "colour without
  // its alpha channel" case: it must be promoted to a valid opaque 8-hex ARGB, not emitted as a
  // 6-char rgb that Excel renders black. A value that is neither 6 nor 8 hex digits is a programming
  // error and must be rejected, never written as a colour Excel silently renders black.
  argbNormalizationReport() {
    const emittedFgColor = argb => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb}};
      const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
      return (stylesXml.match(/<fgColor rgb="([^"]*)"/) || [null, null])[1];
    };
    let rejectsMalformed = false;
    try {
      emittedFgColor('12345');
    } catch {
      rejectsMalformed = true;
    }
    return {sixHexRgb: emittedFgColor('00FF00'), rejectsMalformed};
  },

  outlinePropertiesRoundtrip() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.outline.summaryBelow = false;
    ws.outline.summaryRight = false;
    ws.getCell('A1').value = 'x';
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const outlinePr = (sheetXml.match(/<outlinePr\b[^>]*\/?>/) || [''])[0];
    const reload = readXlsx(buffer);
    const outline = reload.getWorksheet('S').outline;
    return {
      outlinePrEmitted: /summaryBelow="0"/.test(outlinePr) && /summaryRight="0"/.test(outlinePr),
      reReadSummaryBelow: outline.summaryBelow ?? null,
      reReadSummaryRight: outline.summaryRight ?? null,
    };
  },

  // Author a bold cell, then rewrite the emitted <b/> flag to each explicit form and report how
  // the reader reads bold back → { bareTag, valOne, valZero }. A boolean font flag's `val` governs:
  // a bare tag or val="1" is on, val="0" is off — presence alone must not force true.
  fontExplicitFalseBoldReport() {
    const readBoldWith = tag => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').font = {bold: true};
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': xml => xml.replace(/<b ?\/>/, tag),
      });
      const font = reloaded.getWorksheet('S').getCell('A1').font;
      return !!(font && font.bold);
    };
    return {
      bareTag: readBoldWith('<b/>'),
      valOne: readBoldWith('<b val="1"/>'),
      valZero: readBoldWith('<b val="0"/>'),
    };
  },

  // Author cells with italic/strike/underline on, rewrite each flag to its explicit-off form, and
  // report what the reader reads back → { italic, strike, underline }. val="0" turns a boolean flag
  // off; <u val="none"/> is the ABSENCE of an underline, so it must read back falsy — never the
  // truthy string "none".
  fontExplicitOffFlagsReport() {
    const readWith = (baseFont, tagRe, tag, field) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').font = baseFont;
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': xml => xml.replace(tagRe, tag),
      });
      const font = reloaded.getWorksheet('S').getCell('A1').font || {};
      return font[field] ?? null;
    };
    return {
      italic: readWith({italic: true}, /<i ?\/>/, '<i val="0"/>', 'italic'),
      strike: readWith({strike: true}, /<strike ?\/>/, '<strike val="0"/>', 'strike'),
      underline: readWith({underline: true}, /<u ?\/>/, '<u val="none"/>', 'underline'),
    };
  },

  // Inject an xf whose alignment element carries only an explicit-false boolean (wrapText="0" /
  // shrinkToFit="0"), point A1 at it, and report the alignment the reader surfaces → { wrapTextZero,
  // shrinkZero }. An all-false alignment carries no information and must read back as no alignment
  // at all — the raw "0" is a truthy JS string, so a reader guarding on the raw value rather than the
  // parsed boolean would wrongly report a present alignment.
  alignmentFalseBooleanReport() {
    const readWithAlignment = alignAttr => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      let injectedIndex = -1;
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': xml => {
          const count = Number(xml.match(/<cellXfs count="(\d+)">/)[1]);
          injectedIndex = count;
          return xml
            .replace(/<cellXfs count="\d+">/, `<cellXfs count="${count + 1}">`)
            .replace(
              /<\/cellXfs>/,
              `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">` +
                `<alignment ${alignAttr}/></xf></cellXfs>`
            );
        },
        'xl/worksheets/sheet1.xml': xml => xml.replace(/<c r="A1"([^>]*)>/, `<c r="A1" s="${injectedIndex}"$1>`),
      });
      const alignment = reloaded.getWorksheet('S').getCell('A1').alignment;
      return alignment ? JSON.parse(JSON.stringify(alignment)) : null;
    };
    return {
      wrapTextZero: readWithAlignment('wrapText="0"'),
      shrinkZero: readWithAlignment('shrinkToFit="0"'),
    };
  },

  // Read a real fixture `.xlsx` and report only whether it loaded, any error, its sheet names, and
  // a couple of core properties → { ok, error, sheetNames, lastModifiedBy, creator }. The read error
  // is captured and returned as data (never propagated) so a case asserts on a crash rather than the
  // runner blowing up. Exercises the reader against foreign generators and schema-valid corners Excel
  // never emits (namespace-prefixed roots, a leading BOM, unusual part order, missing optional parts).
  readFixtureReport(rel) {
    try {
      const wb = readFixture(rel);
      return {
        ok: true,
        error: null,
        sheetNames: wb.worksheets.map(s => s.name),
        lastModifiedBy: wb.properties.lastModifiedBy ?? null,
        creator: wb.properties.creator ?? null,
      };
    } catch (e) {
      return {ok: false, error: String((e && e.message) || e), sheetNames: null};
    }
  },

  // Read a fixture and report its defined names as { <name>: [refersTo…] }, mirroring the oracle.
  // The model retains every name as its own entry rather than keying by name, so two same-named
  // names scoped to different sheets both survive — the scope collision that drops one on the
  // oracle's name-keyed reader.
  readFixtureDefinedNames(rel) {
    const wb = readFixture(rel);
    const names = {};
    for (const dn of wb.definedNames) (names[dn.name] ||= []).push(dn.refersTo);
    for (const k of Object.keys(names)) names[k].sort();
    return {names, count: Object.keys(names).length, modelCount: wb.definedNames.length};
  },

  // Read a real fixture `.xlsx` and report the fill and font colour the reader surfaces for each
  // requested `<sheet>!<address>` cell → { [key]: { fill, fontColor } | null }. Mirrors the oracle:
  // a solid-pattern fill's visible colour lives on fgColor while bgColor is the automatic indexed
  // placeholder, and the font colour is a wholly separate facet — the two are never conflated.
  // Read a real fixture `.xlsx` and report each requested cell's observable type, value, number
  // format, and note → { <addr>: {type, value, numFmt, note} | null }, on the first sheet. Mirrors
  // the oracle: a date-formatted numeric serial surfaces as a Date (value { date: iso }), not a raw
  // number, honouring the 1900 date-system leap-year quirk. `type` is a stable label.
  readFixtureCells(rel, cells = []) {
    const wb = readFixture(rel);
    const sheet = wb.worksheets[0];
    const out = {};
    for (const addr of cells) {
      const cell = sheet ? sheet.getCell(addr) : null;
      out[addr] = cell
        ? {
            type: cell.type,
            value: normalizeStreamValue(cell.value),
            numFmt: cell.numFmt ?? null,
            note: cell.note !== undefined ? cell.note : undefined,
          }
        : null;
    }
    return out;
  },

  // ── Streaming reader (readWorkbookStream) ──────────────────────────────────────────────────────
  // The corpus's streaming-read contract, bound to the rewrite's generator-based reader. Each method
  // mirrors its oracle sibling in workbook-io.mjs; where a case compares the streaming path to the
  // eager one, BOTH come from the rewrite, so the assertion checks that streaming and buffered reads
  // agree cell-for-cell. The rewrite's reader is a synchronous generator, so the "without race" and
  // "chunk boundary" hazards the ExcelJS stream reader faces are structurally absent.

  // Read a fixture both eagerly and through the streaming reader, reporting the sheet names each
  // surfaces → { eager, streaming }. The streaming reader joins each worksheet part to the
  // workbook-level declaration, so it exposes the real names, not positional placeholders.
  streamVsEagerSheetNames(rel) {
    const eager = readFixture(rel).worksheets.map(s => s.name);
    const streaming = [...readWorkbookStream(fixtureBytes(rel))].map(s => s.name);
    return {eager, streaming};
  },

  // Report the first sheet's populated row numbers from both paths → { eager, streaming }. Both skip
  // fully-empty rows (the eager `includeEmpty:false` intent) so a gap between data rows is preserved
  // as a jump in the numbers, never resequenced.
  streamVsEagerRowNumbers(rel) {
    const es = readFixture(rel).worksheets[0];
    const eager = [];
    if (es) for (const row of es.rows()) if (row.cells.length) eager.push(row.number);
    const streaming = [];
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows()) if (row.cells.length) streaming.push(row.number);
      break; // first worksheet only
    }
    return {eager, streaming};
  },

  // Report each populated first-sheet row's { number, hidden } from both paths → { eager, streaming }.
  // The streaming reader must surface a row's hidden flag (in the string form "true"/"false" some
  // generators write), agreeing with the eager read rather than reporting every row visible.
  streamVsEagerRowHidden(rel) {
    const es = readFixture(rel).worksheets[0];
    const eager = [];
    if (es) for (const row of es.rows()) if (row.cells.length) eager.push({number: row.number, hidden: !!row.properties?.hidden});
    const streaming = [];
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows()) if (row.cells.length) streaming.push({number: row.number, hidden: !!row.hidden});
      break; // first worksheet only
    }
    return {eager, streaming};
  },

  // Write a sheet with a hidden column, then read it eagerly and through the streaming reader,
  // reporting each path's per-column hidden flags → { eager, stream, error }. The streaming reader
  // parses <col hidden> and surfaces it after the rows are drained, matching the eager oracle.
  streamVsEagerColumnHidden() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getColumn(2).hidden = true;
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    s.getCell('C1').value = 'c';
    const buffer = writeXlsx(wb);

    const es = readXlsx(buffer).getWorksheet('S');
    const eager = {col1: !!es.getColumn(1).hidden, col2: !!es.getColumn(2).hidden, col3: !!es.getColumn(3).hidden};

    const stream = {};
    let error = null;
    try {
      for (const sheet of readWorkbookStream(buffer)) {
        for (const _row of sheet.rows()) void _row;
        const hidden = new Set(sheet.hiddenColumns);
        stream.col1 = hidden.has(1);
        stream.col2 = hidden.has(2);
        stream.col3 = hidden.has(3);
        break; // first worksheet only
      }
    } catch (e) {
      error = String((e && e.message) || e);
    }
    return {eager, stream, error};
  },

  // Build a sheet with two merged ranges, then report the merge geometry from both the eager and the
  // streaming path → { eagerMerges, streamedMerges, error }. The streaming reader collects
  // <mergeCells> (which follows <sheetData>) during the same pass and exposes it after the rows.
  streamReadMergesReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'm';
    ws.mergeCells('A1:B2');
    ws.getCell('D1').value = 'n';
    ws.mergeCells('D1:D3');
    const buffer = writeXlsx(wb);

    const eagerMerges = [...readXlsx(buffer).getWorksheet('S').merges].sort();

    let streamedMerges = null;
    let error = null;
    try {
      for (const sheet of readWorkbookStream(buffer)) {
        for (const _row of sheet.rows()) void _row;
        streamedMerges = [...sheet.merges].sort();
        break; // first worksheet only
      }
    } catch (e) {
      error = String((e && e.message) || e);
    }
    return {eagerMerges, streamedMerges, error};
  },

  // Report the 1-based row-values array for the requested rows from both paths → { eager, streamed }.
  // A streamed row indexes exactly as a full-load row does (empty slot at 0, column A at 1), so a
  // caller can switch readers without re-indexing.
  streamVsEagerRowValues(spec, rowNumbers = [1]) {
    const buffer = writeXlsx(buildFrom(spec));
    const wanted = new Set(rowNumbers);

    const es = readXlsx(buffer).worksheets[0];
    const eager = {};
    if (es) for (const row of es.rows()) if (wanted.has(row.number)) eager[row.number] = streamedRowValues(row.cells);
    for (const n of rowNumbers) eager[n] ??= [null];

    const streamed = {};
    for (const sheet of readWorkbookStream(buffer)) {
      for (const row of sheet.rows()) if (wanted.has(row.number)) streamed[row.number] = streamedRowValues(row.cells);
      break; // first worksheet only
    }
    for (const n of rowNumbers) streamed[n] ??= [null];
    return {eager, streamed};
  },

  // Write `count` single-cell worksheets, then stream them back, reporting { written, emitted, error,
  // first, last }. Exercises the reader across far more than 100 sheets and a package whose worksheet
  // parts may precede the workbook part — every sheet must be emitted exactly once.
  streamReadManySheets(count = 180) {
    const wb = new Workbook();
    for (let i = 0; i < count; i++) wb.addWorksheet(`Sheet${i + 1}`).getCell('A1').value = i;
    const buffer = writeXlsx(wb);
    const names = [];
    let error = null;
    try {
      for (const sheet of readWorkbookStream(buffer)) names.push(sheet.name);
    } catch (e) {
      error = String((e && e.message) || e);
    }
    return {written: count, emitted: names.length, error, first: names[0] ?? null, last: names[names.length - 1] ?? null};
  },

  // Write a shared-strings workbook, then read it through the streaming reader once and again
  // concurrently, reporting whether every shared-string cell resolved → { singleComplete, singleLength,
  // concurrentAllComplete, concurrentLengths }. The rewrite's reader is a pure synchronous generator,
  // so concurrent reads cannot race over a shared shared-strings table.
  async streamingSharedStringsRead(rowCount = 20, concurrency = 8) {
    const build = new Workbook();
    const bs = build.addWorksheet('S');
    for (let r = 1; r <= rowCount; r++) {
      bs.getCell(`A${r}`).value = `str${r % 3}`;
      bs.getCell(`B${r}`).value = r;
    }
    const buffer = writeXlsx(build, {useSharedStrings: true});

    const readOne = () => {
      const strings = [];
      for (const sheet of readWorkbookStream(buffer)) {
        for (const row of sheet.rows()) {
          const first = row.cells.find(cell => cell.col === 1);
          strings.push(first ? first.value : undefined);
        }
      }
      return strings;
    };

    const single = readOne();
    const singleComplete = single.length === rowCount && single.every(v => typeof v === 'string');
    const many = await Promise.all(Array.from({length: concurrency}, async () => readOne()));
    const allComplete = many.every(v => v.length === rowCount && v.every(x => typeof x === 'string'));
    return {
      singleComplete,
      singleLength: single.length,
      concurrentAllComplete: allComplete,
      concurrentLengths: many.map(v => v.length),
    };
  },

  // Stream-read a fixture end-to-end, reporting { ok, error, sheetNames, totalRows } — the read either
  // completes (with every sheet name and the total rows delivered) or its error is captured as data.
  // Locks that the reader tolerates a package whose ZIP places a worksheet part before xl/workbook.xml
  // (the inflate builds a path→bytes map, so entry order is irrelevant).
  streamReadReport(rel) {
    const sheetNames = [];
    let totalRows = 0;
    try {
      for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
        sheetNames.push(sheet.name);
        for (const _row of sheet.rows()) totalRows += 1;
      }
      return {ok: true, error: null, sheetNames, totalRows};
    } catch (e) {
      return {ok: false, error: String((e && e.message) || e), sheetNames, totalRows};
    }
  },

  // Stream-read a fixture's first sheet, reporting each requested cell's { type, value } | null. The
  // type is the model's stable label; a date-formatted numeric cell surfaces as a Date because the
  // streaming reader applies the cell's number format when decoding, exactly as the eager read does.
  streamReadFixture(rel, cells = []) {
    const wanted = new Map(cells.map(a => [a, null]));
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows()) {
        for (const cell of row.cells) {
          if (wanted.has(cell.address)) wanted.set(cell.address, {type: detectValueType(cell.value), value: normalizeStreamValue(cell.value)});
        }
      }
      break; // first worksheet only
    }
    const out = {};
    for (const [k, v] of wanted) out[k] = v;
    return out;
  },

  // Write a spec, then read the requested cells through both paths → { streamed, eager }. Proves the
  // streaming reader returns multi-byte UTF-8 text (CJK, emoji) byte-exact and identical to the eager
  // read — the whole-package inflate decodes UTF-8 as one unit, so no character is split.
  streamReadSpec(spec, cells = []) {
    const buffer = writeXlsx(buildFrom(spec));
    const wanted = new Set(cells);
    const streamed = {};
    for (const sheet of readWorkbookStream(buffer)) {
      for (const row of sheet.rows()) {
        for (const cell of row.cells) if (wanted.has(cell.address)) streamed[cell.address] = normalizeStreamValue(cell.value);
      }
      break; // first worksheet only
    }
    const es = readXlsx(buffer).worksheets[0];
    const eager = {};
    for (const ref of cells) eager[ref] = normalizeStreamValue(es ? es.getCell(ref).value : null);
    return {streamed, eager};
  },

  // Give one cell a fill and another a border but NO value, leave a third cell entirely untouched,
  // round-trip, and report each → { filledArgb, filledValue, borderedStyle, borderedValue, untouched }.
  // A formatted-but-empty cell is a real cell Excel keeps: its style must survive the write and the
  // cell must stay value-less, while a cell with neither value nor style must not be fabricated.
  styledEmptyCellReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'anchor';
    sheet.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
    sheet.getCell('C3').border = {top: {style: 'thin', color: {argb: 'FF000000'}}};
    sheet.getCell('D4'); // materialised but never given a value or style
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S');
    const filled = back.getCell('B2');
    const bordered = back.getCell('C3');
    return {
      filledArgb: filled.fill && filled.fill.fgColor ? filled.fill.fgColor.argb : null,
      filledValue: filled.value,
      borderedStyle: bordered.border && bordered.border.top ? bordered.border.top.style : null,
      borderedValue: bordered.value,
      untouched: !!(back.getCell('D4').fill || back.getCell('D4').value),
    };
  },

  readFixtureCellStyles(rel, cells = []) {
    const wb = readFixture(rel);
    const out = {};
    for (const key of cells) {
      const [sheetName, addr] = key.split('!');
      const sheet = wb.getWorksheet(sheetName);
      const cell = sheet ? sheet.getCell(addr) : null;
      out[key] = cell
        ? {
            fill: cell.fill && cell.fill.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
            fontColor: cell.font && cell.font.color ? JSON.parse(JSON.stringify(cell.font.color)) : null,
          }
        : null;
    }
    return out;
  },

  // Read a real fixture, write it straight back out, read the result, and report whether any
  // cell's visible fill colour or border-edge colours changed across that no-op round-trip →
  // { checked, fillMismatches, borderMismatches, fillSample, borderSample }. Theme+tint and
  // indexed-palette references must survive verbatim so the sheet renders identically after a
  // pure open-then-save. Colour comparison is key-order-insensitive.
  roundtripFixtureColorFidelity(rel) {
    const before = readFixture(rel);
    const after = readXlsx(writeXlsx(before));

    const realFill = cell => (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern !== 'none' ? cell.fill : null);
    const borderColors = cell => {
      if (!cell.border) return null;
      const edges = {};
      for (const edge of ['top', 'left', 'right', 'bottom']) {
        if (cell.border[edge] && cell.border[edge].color) edges[edge] = cell.border[edge].color;
      }
      return Object.keys(edges).length ? edges : null;
    };
    const stableSort = v => {
      if (Array.isArray(v)) return v.map(stableSort);
      if (v && typeof v === 'object') {
        const sorted = {};
        for (const k of Object.keys(v).sort()) sorted[k] = stableSort(v[k]);
        return sorted;
      }
      return v;
    };
    const norm = v => JSON.stringify(stableSort(v ?? null));

    let checked = 0;
    let fillMismatches = 0;
    let borderMismatches = 0;
    let fillSample = null;
    let borderSample = null;
    for (const sheet of before.worksheets) {
      const other = after.getWorksheet(sheet.name);
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          if (!realFill(cell) && !borderColors(cell)) continue;
          checked += 1;
          const oc = other ? other.getCell(cell.address) : null;
          const bf = norm(realFill(cell));
          const af = oc ? norm(realFill(oc)) : '(missing)';
          if (bf !== af) {
            fillMismatches += 1;
            if (!fillSample) fillSample = {cell: `${sheet.name}!${cell.address}`, before: bf, after: af};
          }
          const bb = norm(borderColors(cell));
          const ab = oc ? norm(borderColors(oc)) : '(missing)';
          if (bb !== ab) {
            borderMismatches += 1;
            if (!borderSample) borderSample = {cell: `${sheet.name}!${cell.address}`, before: bb, after: ab};
          }
        }
      }
    }
    return {checked, fillMismatches, borderMismatches, fillSample, borderSample};
  },

  // Read a real styled template, write it straight back out, read the result, and report whether
  // its sheet names, custom column widths, and per-cell fill/font/numFmt/alignment/border survived
  // that no-op read→write→read → { sheetNames(Before), columns(Before), styleSurvival }. This is the
  // mainstream "open a styled template, fill it in, save it" path, which must be format-preserving.
  // Style comparison is key-order-insensitive so a case asserts on content survival, not
  // serialization incidentals. In the rewrite's model a column stores a width only when it is a
  // custom width, so "has a width" is exactly "is a custom width".
  roundtripFixture(rel) {
    const before = readFixture(rel);
    const after = readXlsx(writeXlsx(before));

    const stableSort = v => {
      if (Array.isArray(v)) return v.map(stableSort);
      if (v && typeof v === 'object') {
        const sorted = {};
        for (const k of Object.keys(v).sort()) sorted[k] = stableSort(v[k]);
        return sorted;
      }
      return v;
    };
    const hasStyle = cell =>
      !!(cell.numFmt || (cell.fill && cell.fill.type) || cell.font || cell.alignment || cell.border);
    const styleKey = cell =>
      JSON.stringify(
        stableSort({
          numFmt: cell.numFmt || null,
          fill: cell.fill && cell.fill.type ? cell.fill : null,
          font: cell.font || null,
          alignment: cell.alignment || null,
          border: cell.border || null,
        })
      );
    const columnsWithWidth = wb => {
      const out = {};
      for (const sheet of wb.worksheets) {
        const cols = {};
        for (const {index, properties} of sheet.columns()) {
          if (properties.width !== undefined) cols[index] = {width: properties.width, customWidth: true};
        }
        out[sheet.name] = cols;
      }
      return out;
    };

    let checked = 0;
    let mismatches = 0;
    let sample = null;
    for (const sheet of before.worksheets) {
      const other = after.getWorksheet(sheet.name);
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          // Resolve both sides through getCell so a merged-range slave redirects to its master on
          // each — comparing the row-iterated slave (its own style) against getCell (the master) would
          // report a phantom drift that is only an access asymmetry, not a lost style.
          const beforeCell = sheet.getCell(cell.address);
          if (!hasStyle(beforeCell)) continue;
          checked += 1;
          const beforeKey = styleKey(beforeCell);
          const afterKey = other ? styleKey(other.getCell(cell.address)) : '(sheet missing)';
          if (beforeKey !== afterKey) {
            mismatches += 1;
            if (!sample) sample = {cell: `${sheet.name}!${cell.address}`, before: beforeKey, after: afterKey};
          }
        }
      }
    }

    return {
      sheetNamesBefore: before.worksheets.map(s => s.name),
      sheetNames: after.worksheets.map(s => s.name),
      columnsBefore: columnsWithWidth(before),
      columns: columnsWithWidth(after),
      styleSurvival: {checked, mismatches, sample},
    };
  },

  // Give one column a right border and later columns only a width, then round-trip and report
  // each cell's right border → { a1, b1, c1 }. A column's border is a default for its own cells,
  // so the declaring column's cell carries it while columns without a style of their own get
  // nothing — column styles are independent, not bled into subsequent columns.
  columnBorderScopedReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getColumn(1).border = {right: {style: 'thin', color: {argb: 'FF000000'}}};
    sheet.getColumn(2).width = 10;
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    sheet.getCell('C1').value = 'c';
    const s = readXlsx(writeXlsx(wb)).getWorksheet('S');
    const rightBorder = ref => {
      const b = s.getCell(ref).border;
      return !!(b && b.right && b.right.style);
    };
    return {a1: rightBorder('A1'), b1: rightBorder('B1'), c1: rightBorder('C1')};
  },

  // Build → write → read back through the rewrite's own reader, then normalize to the
  // same JSON model current.mjs reports, so every write→read round-trip case runs
  // unchanged. Facets the writer/reader do not materialize yet come back empty/null;
  // the writer's feature-gate keeps a case whose spec needs those from ever running here.
  roundtripWorkbook(spec) {
    const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
    const sheets = {};
    for (const s of spec.sheets || []) {
      const sheet = reloaded.getWorksheet(s.name);
      if (!sheet) {
        sheets[s.name] = null;
        continue;
      }
      const cells = {};
      for (const c of s.cells || []) cells[c.ref] = normalizeRewriteCell(sheet.getCell(c.ref));
      const columns = {};
      for (const col of s.columns || []) {
        const p = sheet.getColumn(col.index);
        columns[col.index] = {width: p.width ?? null, hidden: !!p.hidden, numFmt: p.numFmt ?? null};
      }
      const rows = {};
      for (const row of s.rows || []) {
        const p = sheet.getRow(row.index);
        rows[row.index] = {height: p.height ?? null, hidden: !!p.hidden};
      }
      const margins = Object.keys(sheet.pageMargins).length > 0 ? {...sheet.pageMargins} : null;
      const ps = sheet.pageSetup;
      sheets[s.name] = {
        cells,
        columns,
        rows,
        margins,
        pageSetup: {
          fitToPage: !!ps.fitToPage,
          fitToWidth: ps.fitToWidth ?? null,
          fitToHeight: ps.fitToHeight ?? null,
          scale: ps.scale ?? null,
          paperSize: ps.paperSize ?? null,
        },
        autoFilter: sheet.autoFilter?.ref ?? null,
        merges: [...sheet.merges],
        rowCount: sheet.rowCount,
        actualRowCount: sheet.actualRowCount,
      };
    }
    const props = reloaded.properties;
    const definedNames = {};
    for (const dn of reloaded.definedNames) (definedNames[dn.name] ||= []).push(dn.refersTo);
    for (const k of Object.keys(definedNames)) definedNames[k].sort();
    return {
      properties: {
        creator: props.creator ?? null,
        lastModifiedBy: props.lastModifiedBy ?? null,
        created: isoOrNull(props.created),
        modified: isoOrNull(props.modified),
      },
      sheets,
      definedNames,
    };
  },

  // Write a rich-text cell, then read it back, and report how its runs serialized and survived →
  // { emptyTextRunInXml, runCount, runs: [{text, bold, italic, underline}] }. Mirrors the oracle. A
  // zero-length run must never emit an empty <t> (Excel flags it corrupt); the surviving runs keep
  // their text and per-run formatting. The rewrite writes rich text inline, so both the empty-<t>
  // scan and the read-back target the worksheet XML (there is no shared-strings part).
  async richTextRoundtripReport(runs) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = {richText: runs};
    const buffer = writeXlsx(wb);
    const parts = partMapOf(buffer);
    const xml = parts['xl/sharedStrings.xml'] || parts['xl/worksheets/sheet1.xml'] || '';
    const emptyTextRunInXml = /<(?:\w+:)?t\b[^>]*\/>|<(?:\w+:)?t\b[^>]*><\/(?:\w+:)?t>/.test(xml);
    const value = readXlsx(buffer).getWorksheet('S').getCell('A1').value;
    const readRuns = value && Array.isArray(value.richText) ? value.richText : [];
    return {
      emptyTextRunInXml,
      runCount: readRuns.length,
      runs: readRuns.map(r => ({
        text: r.text ?? null,
        bold: r.font ? r.font.bold ?? false : false,
        italic: r.font ? r.font.italic ?? false : false,
        underline: r.font ? r.font.underline ?? false : false,
      })),
    };
  },

  // Read a fixture and report every hyperlink cell as { <addr>: { hyperlink, text } }, with a rich
  // display label flattened to its concatenated text — for asserting a foreign file's links (and the
  // rejoining of an external URL's fragment carried in the location attribute) are read faithfully.
  async readFixtureHyperlinks(rel) {
    const flatten = t =>
      t == null ? null : typeof t === 'string' ? t : Array.isArray(t.richText) ? t.richText.map(r => r.text).join('') : t;
    const sheet = readFixture(rel).worksheets[0];
    const out = {};
    if (sheet) {
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          const v = cell.value;
          if (v && typeof v === 'object' && 'hyperlink' in v) {
            out[cell.address] = {hyperlink: v.hyperlink ?? null, text: flatten(v.text)};
          }
        }
      }
    }
    return out;
  },

  // Read a fixture and report each covered cell's data validation → { cells: { 'Sheet!Ref': rule },
  // count }. A rule authored over a multi-cell range must be reported on EVERY covered cell (the
  // range form is resolved per cell), and a reference/name operand must survive as its string.
  readFixtureValidations(rel) {
    const wb = readFixture(rel);
    const cells = {};
    for (const sheet of wb.worksheets) {
      for (const {sqref} of sheet.dataValidations) {
        for (const ref of expandSqref(sqref)) {
          const dv = sheet.dataValidationAt(ref);
          if (dv) cells[`${sheet.name}!${ref}`] = JSON.parse(JSON.stringify(dv));
        }
      }
    }
    return {cells, count: Object.keys(cells).length};
  },

  // Read a fixture and report each sheet's data-validation rules, de-duplicated by content with a
  // per-rule coverage count → { sheets: { name: { rules: [{rule, coverageCount}], ruleCount } } }.
  // Reads the worksheet overlay (not populated cells), so a rule over an empty range is still seen,
  // and surfaces a reference source (a defined name, a cross-sheet range) as its verbatim formula
  // text rather than "[object Object]".
  readFixtureValidationRules(rel) {
    const wb = readFixture(rel);
    const sheets = {};
    for (const sheet of wb.worksheets) {
      const byContent = new Map();
      for (const {sqref, rule} of sheet.dataValidations) {
        const key = JSON.stringify(rule);
        const entry = byContent.get(key) || {rule: JSON.parse(key), coverageCount: 0};
        entry.coverageCount += expandSqref(sqref).length;
        byContent.set(key, entry);
      }
      sheets[sheet.name] = {rules: [...byContent.values()], ruleCount: byContent.size};
    }
    return {sheets};
  },

  // Read a fixture, write it back, and report the data-validation facts of the *re-serialized*
  // package — both the standard `<dataValidation>` entries and the extended `<x14:dataValidation>`
  // form (2009 extension schema, carried in `<extLst>`, used for cross-sheet / whole-column list
  // sources). Lets a case assert a template's validation survives a read→write round-trip rather than
  // being silently dropped because only the standard form was understood.
  roundtripFixtureValidationXml(rel) {
    const parts = partMapOf(writeXlsx(readFixture(rel)));
    const sheetParts = Object.keys(parts)
      .filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p))
      .sort();

    const sheets = {};
    let totalStandard = 0;
    let totalExt = 0;
    for (const p of sheetParts) {
      const xml = parts[p] || '';
      // `[ >]` after the tag name separates an individual entry from its `<dataValidations>` /
      // `<x14:dataValidations>` container (whose next char is `s`).
      const standardCount = [...xml.matchAll(/<dataValidation[ >]/g)].length;
      const extCount = [...xml.matchAll(/<x14:dataValidation[ >]/g)].length;
      const extSqrefs = [...xml.matchAll(/<xm:sqref>([^<]*)<\/xm:sqref>/g)].map(m => m[1]);
      const standardRules = [
        ...xml.matchAll(/<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g),
      ].map(m => {
        const a = attrsOf('<x ' + m[1] + '>');
        const f1 = (m[2].match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
        return {
          type: a.type ?? null,
          sqref: a.sqref ?? null,
          errorTitle: a.errorTitle ?? null,
          error: a.error ?? null,
          formula1: f1 == null ? null : f1.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'),
        };
      });
      sheets[p] = {standardCount, extCount, extSqrefs, standardRules, hasExtLst: /<extLst\b/.test(xml)};
      totalStandard += standardCount;
      totalExt += extCount;
    }
    return {sheets, totalStandard, totalExt, totalValidations: totalStandard + totalExt};
  },

  // Attach a list validation whose source formula is supplied (possibly with a leading '='), write,
  // and report the serialized formula1 text → { formula1, hasLeadingEquals }. OOXML formula1 carries
  // no leading '='; the writer must strip exactly one so the app applies the validation immediately.
  dvFormulaLeadingEquals(formula = '=$AA$1:$AA$2') {
    const wb = new Workbook();
    wb.addWorksheet('S').addDataValidation('A1', {type: 'list', allowBlank: true, formulae: [formula]});
    const sheetXml = partMapOf(writeXlsx(wb))['xl/worksheets/sheet1.xml'] || '';
    const formula1 = (sheetXml.match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
    return {formula1, hasLeadingEquals: formula1 != null && formula1.startsWith('=')};
  },

  // Attach one validation over a whole range, write, and report the serialized facts → { writeOk,
  // writeError, sqrefs, count, reloadOk }. A range-form validation must emit exactly ONE
  // dataValidation whose sqref is the requested range, not one entry per covered cell.
  roundtripRangeValidation({range, type = 'list', formula = '"a,b,c"'} = {}) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    let buffer;
    try {
      sheet.addDataValidation(range, {type, allowBlank: true, formulae: [formula]});
      buffer = writeXlsx(wb);
    } catch (e) {
      return {writeOk: false, writeError: String((e && e.message) || e), sqrefs: [], count: 0, reloadOk: null};
    }
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const sqrefs = [...xml.matchAll(/<dataValidation\b[^>]*sqref="([^"]*)"/g)].map(m => m[1]);
    const count = [...xml.matchAll(/<dataValidation[ >]/g)].length;
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {writeOk: true, writeError: null, sqrefs, count, reloadOk};
  },

  // Author list validations on a 'Main' sheet from the two source forms an author uses — an inline
  // quoted literal ("Male,Female") and a cross-sheet range reference (Levels!$A$2:$A$9999) — write,
  // read back, and report both the per-cell rule the reader hands back and the serialized
  // `<dataValidations>` facts (count, well-formedness, the verbatim formula1 texts). Lets a case
  // assert BOTH forms survive a write→read round-trip and that inline lists stay quoted while range
  // references stay unquoted, without the case knowing how validations are shaped internally.
  authorListValidations(validations = []) {
    const wb = new Workbook();
    const main = wb.addWorksheet('Main');
    const levels = wb.addWorksheet('Levels');
    levels.getCell('A2').value = 'X';
    for (const v of validations) {
      const rule = {type: 'list', allowBlank: v.allowBlank !== false, formulae: [v.formula]};
      if (v.error !== undefined) {
        rule.showErrorMessage = true;
        rule.error = v.error;
      }
      main.addDataValidation(v.ref, rule);
    }
    const buffer = writeXlsx(wb);

    const reread = readXlsx(buffer).getWorksheet('Main');
    const readBack = {};
    for (const v of validations) {
      const dv = reread?.dataValidationAt(v.ref);
      readBack[v.ref] = dv ? {type: dv.type, formulae: dv.formulae ?? null} : null;
    }

    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const block = (xml.match(/<dataValidations[\s\S]*?<\/dataValidations>/) || [])[0] || '';
    return {
      readBack,
      xml: {
        count: [...xml.matchAll(/<dataValidation[ >]/g)].length,
        // Cheap structural check: a strict consumer chokes on a raw & that is not an entity.
        wellFormed: !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(block),
        formula1: [...block.matchAll(/<formula1>([\s\S]*?)<\/formula1>/g)].map(m =>
          m[1].replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
        ),
      },
    };
  },

  // Build a workbook with one internal ('#'-prefixed) hyperlink, write it, and report how the link
  // serialized → { writeOk, hasLocation, location, hasExternalRel, hasRid, reloadOk }. An internal
  // target must ride in a `location` attribute with no external-mode relationship, and the package
  // must reload.
  async internalHyperlinkReport(target = "#'Target'!A1") {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Main');
    wb.addWorksheet('Target');
    sheet.getCell('A1').value = {text: 'go', hyperlink: target};
    let buffer;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      return {writeOk: false, writeError: String((e && e.message) || e)};
    }
    const parts = partMapOf(buffer);
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const relsXml = parts['xl/worksheets/_rels/sheet1.xml.rels'] || '';
    const a = attrsOf((sheetXml.match(/<hyperlink\b[^>]*\/?>/) || [''])[0]);
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      writeOk: true,
      hasLocation: a.location != null,
      location: a.location ?? null,
      hasExternalRel: /TargetMode="External"/.test(relsXml),
      hasRid: a['r:id'] != null,
      reloadOk,
    };
  },

  // Build a workbook with an internal '#Sheet2!A1' hyperlink (plus a tooltip), write it, and report
  // the serialized distinctions → { hasWorksheetRels, hyperlinkHasRid, hyperlinkLocation,
  // relTargetMode, reReadHyperlink }. The internal form must carry a location and NO external
  // relationship, and the target must survive a reload.
  async internalHyperlinkSerializationReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    wb.addWorksheet('Sheet2');
    ws.getCell('A1').value = {text: 'go', hyperlink: '#Sheet2!A1', tooltip: 'tt'};
    const buffer = writeXlsx(wb);
    const parts = partMapOf(buffer);
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const relsXml = parts['xl/worksheets/_rels/sheet1.xml.rels'] || '';
    const hyperlinkEl = (sheetXml.match(/<hyperlink\b[^>]*\/?>/) || [''])[0];
    const relEl = (relsXml.match(/<Relationship\b[^>]*hyperlink[^>]*\/?>/) || [''])[0];
    const reReadHyperlink = readXlsx(buffer).getWorksheet('Sheet1').getCell('A1').value.hyperlink ?? null;
    return {
      hasWorksheetRels: /Type="[^"]*\/hyperlink"/.test(relsXml),
      hyperlinkHasRid: /r:id="/.test(hyperlinkEl),
      hyperlinkLocation: (hyperlinkEl.match(/location="([^"]*)"/) || [null, null])[1],
      relTargetMode: (relEl.match(/TargetMode="([^"]*)"/) || [null, null])[1],
      reReadHyperlink,
    };
  },

  // Read a fixture, extract its column widths and pageSetup, write it straight back, re-read, and
  // report the same facts → { source, rewritten }. A faithful no-op round-trip must reproduce every
  // fractional column width and the print-scaling attributes the real file carries.
  roundtripFixtureStyleFacts(rel) {
    const facts = workbook => {
      const sheet = workbook.worksheets[0];
      const ps = sheet ? sheet.pageSetup : {};
      // Differential styles are preserved as verbatim `<dxf>` fragments; a rule's number format is
      // whatever formatCode the fragment carries, so a coerced "[object Object]" can never appear.
      const dxfs = workbook.differentialStyles;
      const dxfFormatCodes = dxfs.flatMap(f =>
        [...f.matchAll(/formatCode="([^"]*)"/g)].map(m => m[1])
      );
      return {
        columnWidths: sheet
          ? [...sheet.columns()].map(c => c.properties.width).filter(w => w !== undefined)
          : [],
        pageSetup: {
          scale: ps.scale ?? null,
          fitToWidth: ps.fitToWidth ?? null,
          fitToHeight: ps.fitToHeight ?? null,
          pageOrder: ps.pageOrder ?? null,
          orientation: ps.orientation ?? null,
          paperSize: ps.paperSize ?? null,
        },
        dxfCount: dxfs.length,
        dxfFormatCodes,
      };
    };
    const before = readFixture(rel);
    const source = facts(before);
    const after = readXlsx(writeXlsx(before));
    return {source, rewritten: facts(after)};
  },

  // Author a conditional-formatting rule, write it, and report the emitted CF XML facts plus what the
  // reader surfaces on reload → { writeOk, writeError, xml:{blockCount, sqrefs, ruleCount, hasDataBar,
  // cfvoCount, hasColor, wellFormed}, reload:{type, color, gradient, cfvo} }.
  authorConditionalFormatting(cf) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    // Populate the ref's first column so the rule binds to real cells.
    const rows = Number((cf.ref.match(/(\d+)\s*$/) || [])[1] || 3);
    for (let r = 1; r <= rows; r++) sheet.getCell(`A${r}`).value = r / rows;
    let buffer;
    try {
      sheet.addConditionalFormatting(cf);
      buffer = writeXlsx(workbook);
    } catch (e) {
      return {writeOk: false, writeError: String((e && e.message) || e), xml: null, reload: null};
    }
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const cfBlock = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [''])[0];
    const dataBar = (cfBlock.match(/<dataBar\b[\s\S]*?<\/dataBar>|<dataBar\b[^>]*\/>/) || [''])[0];
    const rule = readXlsx(buffer).getWorksheet('S')?.conditionalFormattings?.[0]?.rules?.[0] ?? null;
    return {
      writeOk: true,
      writeError: null,
      xml: {
        blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
        sqrefs: [...xml.matchAll(/<conditionalFormatting\b[^>]*sqref="([^"]*)"/g)].map(m => m[1]),
        ruleCount: [...cfBlock.matchAll(/<cfRule\b/g)].length,
        hasDataBar: /<dataBar\b/.test(cfBlock),
        cfvoCount: [...dataBar.matchAll(/<cfvo\b/g)].length,
        hasColor: /<color\b/.test(dataBar),
        wellFormed: cfBlock ? !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(cfBlock) : false,
      },
      reload: rule
        ? {
            type: rule.type ?? null,
            color: rule.color ? rule.color.argb ?? null : null,
            gradient: rule.gradient ?? null,
            cfvo: (rule.cfvo || []).map(v => ({type: v.type ?? null, value: v.value ?? null})),
          }
        : null,
    };
  },

  // Apply a stopIfTrue rule, write, reload → { xmlHasStopIfTrue, reloadStopIfTrue }.
  conditionalFormattingStopIfTrue() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 5;
    sheet.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'cellIs',
          operator: 'greaterThan',
          formulae: [3],
          stopIfTrue: true,
          style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFFF0000'}}},
        },
      ],
    });
    const buffer = writeXlsx(workbook);
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const rule = readXlsx(buffer).getWorksheet('S')?.conditionalFormattings?.[0]?.rules?.[0];
    return {
      xmlHasStopIfTrue: /stopIfTrue="1"/.test(xml),
      reloadStopIfTrue: rule ? rule.stopIfTrue ?? false : null,
    };
  },

  // Read a fixture's first-sheet conditional-formatting facts, write it back, and report the same
  // before/after → { source, rewritten } each { blockCount, rules:[{type, dxfId, priority}] }.
  roundtripFixtureConditionalFormatting(rel) {
    const cfFacts = xml => ({
      blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
      rules: [...xml.matchAll(/<cfRule\b([^>]*?)\/?>/g)].map(m => {
        const a = attrsOf('<x ' + m[1] + '>');
        return {type: a.type ?? null, dxfId: a.dxfId ?? null, priority: a.priority ?? null};
      }),
    });
    const srcParts = partMapOf(fs.readFileSync(path.join(FIXTURES_ROOT, rel)));
    const srcName = Object.keys(srcParts).find(n => /sheet1\.xml$/.test(n));
    const source = cfFacts(srcName ? srcParts[srcName] : '');
    const outXml = partMapOf(writeXlsx(readFixture(rel)))['xl/worksheets/sheet1.xml'] || '';
    return {source, rewritten: cfFacts(outXml)};
  },

  // Load a fixture and try to write it back → { loadOk, loadError, writeOk, writeError, sheetNames } —
  // for asserting a foreign construct round-trips without the writer crashing.
  roundtripFixtureWriteReport(rel) {
    let workbook;
    try {
      workbook = readFixture(rel);
    } catch (e) {
      return {loadOk: false, loadError: String((e && e.message) || e), writeOk: false, writeError: null, sheetNames: []};
    }
    let writeOk = false;
    let writeError = null;
    try {
      writeXlsx(workbook);
      writeOk = true;
    } catch (e) {
      writeError = String((e && e.message) || e);
    }
    return {loadOk: true, loadError: null, writeOk, writeError, sheetNames: workbook.worksheets.map(w => w.name)};
  },

  // Merge a horizontal span with a value + alignment on the anchor, write, then read back →
  // { mergeCount, merges, populatedCoveredCells, anchorValue, anchorAlignment }. A clean merge
  // declares the range exactly once and emits a value only on the anchor, so the covered cells
  // carry no conflicting <v> — the shape that opens without Excel's repair prompt — and the
  // anchor's value and alignment survive the round-trip.
  mergeCleanReport({anchor = 'B1', range = 'B1:G1', value = 'Group Title'} = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell(anchor);
    cell.value = value;
    cell.alignment = {horizontal: 'center'};
    sheet.mergeCells(range);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const merges = [...sheetXml.matchAll(/<mergeCell\b[^>]*ref="([^"]*)"/g)].map(m => m[1]);
    const {left, right, top, bottom} = decodeRange(range);
    const populatedCoveredCells = [];
    for (let r = top; r <= bottom; r++) {
      for (let c = left; c <= right; c++) {
        const ref = encodeAddress(c, r);
        if (ref === anchor) continue;
        if (new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<v>`).test(sheetXml)) populatedCoveredCells.push(ref);
      }
    }
    const a = readXlsx(buffer).getWorksheet('S').getCell(anchor);
    return {
      mergeCount: merges.length,
      merges,
      populatedCoveredCells,
      anchorValue: a.value ?? null,
      anchorAlignment: a.alignment ? {...a.alignment} : null,
    };
  },

  // Give the top-left/master cell a border (+ numFmt + font), merge it into a region, round-trip,
  // then report the master's border/numFmt/font and the declared merges → for asserting a merge
  // does not strip the style the master needs to render the merged region's outline.
  mergeMasterBorderReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = 'x';
    cell.border = {top: {style: 'thin'}, bottom: {style: 'medium'}};
    cell.numFmt = '0.00';
    cell.font = {bold: true};
    sheet.mergeCells('A1:B2');
    const reread = readXlsx(writeXlsx(workbook)).getWorksheet('S');
    const m = reread.getCell('A1');
    const b = m.border || {};
    return {
      hasTopBorder: !!(b.top && b.top.style),
      hasBottomBorder: !!(b.bottom && b.bottom.style),
      topStyle: b.top ? b.top.style ?? null : null,
      bottomStyle: b.bottom ? b.bottom.style ?? null : null,
      numFmt: m.numFmt ?? null,
      fontBold: !!(m.font && m.font.bold),
      merges: [...reread.merges],
    };
  },

  // Merge a rectangular region (master = top-left), set a value by addressing a NON-master
  // (slave) cell inside it, write, and report which cells carry an independent <v> in the sheet
  // XML, the declared merges, and the re-read master/slave values → for asserting the slave write
  // resolves to the master (only the master carries a value; reading either address returns it).
  mergeSlaveWrite({range = 'A1:B2', slave = 'B2', value = 'slave-write'} = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.mergeCells(range);
    sheet.getCell(slave).value = value;
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    // A cell "carries a value" if its element has value content — a number/bool/formula (<v>),
    // an inline string (<is>), or a formula (<f>). The writer serialises strings as inlineStr,
    // so keying on <v> alone would miss them; an empty covered cell is never emitted at all.
    const cellsWithValue = [...sheetXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>/g)]
      .filter(m => /<(?:v|is|f)\b/.test(m[2]))
      .map(m => m[1]);
    const merges = [...sheetXml.matchAll(/<mergeCell\b[^>]*ref="([^"]*)"/g)].map(m => m[1]);
    const master = range.split(':')[0];
    const s = readXlsx(buffer).getWorksheet('S');
    return {
      cellsWithValue,
      merges,
      masterValue: s.getCell(master).value ?? null,
      slaveValue: s.getCell(slave).value ?? null,
    };
  },

  // Clone a worksheet through its model export/import: build a source sheet with cells and merges,
  // read its model, and assign that model onto a fresh sheet. Reports the merges the source model
  // exposed and the merges the destination carries afterwards → { srcMerges, dstMerges, error }.
  // The historical bug this measures is an asymmetric model contract that dropped merges on import;
  // the rewrite's getter and setter cover the same fields, so the round-trip is lossless.
  copyWorksheetModel({merges = ['A1:C1'], cells = [{ref: 'A1', value: 'merged'}]} = {}) {
    const workbook = new Workbook();
    const src = workbook.addWorksheet('Src');
    for (const c of cells) src.getCell(c.ref).value = c.value;
    for (const m of merges) src.mergeCells(m);
    const dst = workbook.addWorksheet('Dst');

    let error = null;
    let dstMerges = [];
    const srcMerges = [...src.model.merges];
    try {
      dst.model = {...src.model, name: 'Dst'};
      dstMerges = [...dst.model.merges];
    } catch (e) {
      error = String((e && e.message) || e);
    }
    return {srcMerges: srcMerges.sort(), dstMerges: dstMerges.sort(), error};
  },

  mutateWorksheet({cells = [], ops = [], read = [], readStyles = []} = {}) {
    const sheet = new Workbook().addWorksheet('S');
    for (const c of cells) {
      const cell = sheet.getCell(c.ref);
      cell.value = c.value;
      // Optional per-cell style so a case can assert a structural edit carries a cell's style to its
      // shifted position rather than blanking it.
      if (c.font) cell.font = c.font;
      if (c.fill) cell.fill = c.fill;
      if (c.numFmt) cell.numFmt = c.numFmt;
    }

    let error = null;
    try {
      for (const op of ops) {
        const inserts = op.inserts || [];
        if (op.op === 'spliceRows') sheet.spliceRows(op.start, op.count, ...inserts);
        else if (op.op === 'spliceColumns') sheet.spliceColumns(op.start, op.count, ...inserts);
        else if (op.op === 'mergeCells') sheet.mergeCells(op.range);
        else if (op.op === 'insertRow') sheet.insertRow(op.pos, op.value || []);
        else if (op.op === 'duplicateRow') sheet.duplicateRow(op.start, op.count ?? 1, op.insert !== false);
        else throw new Error(`unknown mutation op: ${op.op}`);
      }
    } catch (e) {
      error = String((e && e.message) || e);
    }

    const readCells = {};
    for (const ref of read) readCells[ref] = sheet.getCell(ref).value ?? null;

    // Per-cell style facets after the mutations — for asserting the style a cell carried before a
    // splice still describes the (possibly shifted) cell afterward, rather than being lost.
    const styles = {};
    for (const ref of readStyles) {
      const cell = sheet.getCell(ref);
      styles[ref] = {
        value: cell.value ?? null,
        font: cell.font ? JSON.parse(JSON.stringify(cell.font)) : null,
        fill: cell.fill && cell.fill.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
        numFmt: cell.numFmt ?? null,
      };
    }

    // The last POPULATED row and its column-1 value, derived from the row iterator (ascending, so
    // the final populated row wins) — a delete-splice must leave this on the true last row, never a
    // trailing empty slot.
    let lastRow = null;
    for (const {number, cells: rowCells} of sheet.rows()) {
      if (rowCells.some(c => c.value !== null && c.value !== undefined)) {
        const first = rowCells.find(c => c.col === 1);
        lastRow = {number, value: first?.value ?? null};
      }
    }

    return {
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      cells: readCells,
      styles,
      merges: [...sheet.merges],
      lastRow,
      error,
    };
  },

  // Duplicate a populated row with default args, then merge a range on the copy — for asserting the
  // copy is faithful (values, not empty/NaN) and carries no phantom merge that would reject the merge.
  duplicateRowReport() {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    sheet.getCell('C1').value = 'c';
    let dupError = null;
    try {
      sheet.duplicateRow(1, 1, true);
    } catch (e) {
      dupError = String((e && e.message) || e);
    }
    const val = ref => sheet.getCell(ref).value ?? null;
    const row1 = [val('A1'), val('B1'), val('C1')];
    const row2 = [val('A2'), val('B2'), val('C2')];
    let mergeError = null;
    try {
      sheet.mergeCells('A2:C2');
    } catch (e) {
      mergeError = String((e && e.message) || e);
    }
    return {dupError, mergeError, rowCount: sheet.rowCount, row1, row2, merges: [...sheet.merges]};
  },

  // Insert a row then style a cell of it — for asserting the inserted cells stay mutable (no frozen,
  // "object is not extensible" style object) regardless of the requested style-inheritance mode. The
  // rewrite's copy-on-write style model makes every cell mutable by construction, so the mode is
  // immaterial; it is accepted and ignored.
  insertRowThenStyle(_styleMode = 'i') {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.getCell('A1').font = {bold: true};
    sheet.getCell('A2').value = 'data';
    let error = null;
    let numFmt = null;
    try {
      sheet.insertRow(2, ['inserted']);
      const cell = sheet.getCell('A2');
      cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
      cell.font = {...cell.font, bold: true};
      numFmt = cell.numFmt;
    } catch (e) {
      error = String((e && e.message) || e);
    }
    return {error, numFmt};
  },

  // Attach a cell note and an outline level, insert a row above them, and round-trip through the real
  // writer/reader → { dataShifted, noteFollowsRow, outlineFollowsRow }. Both the note and the outline
  // level must follow their logical row through the insert and survive serialization.
  rowInsertPreservesNoteAndOutline() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'r1';
    ws.getCell('A2').value = 'r2';
    ws.getCell('A2').note = 'mynote';
    ws.getRow(2).outlineLevel = 1;
    ws.insertRow(1, ['new']); // r1 -> row 2, r2 (noted, outlined) -> row 3
    const s = readXlsx(writeXlsx(wb)).getWorksheet('S');
    return {
      dataShifted: s.getCell('A2').value === 'r1' && s.getCell('A3').value === 'r2',
      noteFollowsRow: !!s.getCell('A3').note,
      outlineFollowsRow: s.getRow(3).outlineLevel === 1,
    };
  },

  spliceShiftsRefs() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    // Table occupies A3:B5 (header + 2 rows); image anchored from row 5 (0-based).
    s.addTable({name: 'T', ref: 'A3', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
    const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    s.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
    s.spliceRows(1, 0, ['inserted']); // insert a row at the top → table and image shift down 1

    const parts = partMapOf(writeXlsx(wb));
    const tableXml = parts['xl/tables/table1.xml'] || '';
    const tableRef = ((tableXml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1]) ?? null;
    const drawingXml = parts['xl/drawings/drawing1.xml'] || '';
    const imageFromRow = (drawingXml.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/) || [])[1];

    // Duplicate table column names authored separately — construction disambiguates them into a
    // unique set rather than emitting a corrupt table (the same repair the reader applies on load).
    const w2 = new Workbook();
    const dupTable = w2
      .addWorksheet('S')
      .addTable({name: 'T2', ref: 'A1', columns: [{name: 'Dup'}, {name: 'Dup'}], rowCount: 1});
    writeXlsx(w2);
    const dupColumnNames = dupTable.columns.map(c => c.name);
    const dupColumnNamesUnique =
      new Set(dupColumnNames.map(n => n.toLowerCase())).size === dupColumnNames.length;

    return {
      tableRef,
      imageFromRow: imageFromRow != null ? Number(imageFromRow) : null,
      dupColumnNames,
      dupColumnNamesUnique,
    };
  },

  // Find a table by name across a loaded fixture's sheets and report its column names and data-row
  // count. The reader reconstructs the table from its part, deriving the data-row count from the
  // stored range (height minus the header and totals rows), so a loaded table exposes its rows.
  readFixtureTable(rel, tableName) {
    const wb = readFixture(rel);
    for (const s of wb.worksheets) {
      const table = s.tables.find(t => t.name === tableName);
      if (table) return {found: true, columns: table.columns.map(c => c.name), rowCount: table.options.rowCount};
    }
    return {found: false, columns: null, rowCount: null};
  },

  // Load a fixture and report a named table's column count and names — used to prove a table with a
  // calculated column (a <calculatedColumnFormula> child the reader ignores) does not truncate the
  // column list or crash the read.
  loadFixtureTableColumns(rel, tableName) {
    try {
      const wb = readFixture(rel);
      for (const s of wb.worksheets) {
        const table = s.tables.find(t => t.name === tableName);
        if (table) {
          return {loaded: true, error: null, columnCount: table.columns.length, columnNames: table.columns.map(c => c.name)};
        }
      }
      return {loaded: true, error: null, columnCount: 0, columnNames: []};
    } catch (e) {
      return {loaded: false, error: String((e && e.message) || e), columnCount: null, columnNames: null};
    }
  },

  // Build a table-bearing spec, report the full ref written into each table part, then read the
  // package back and re-write it, reporting the ref and well-formedness of each re-emitted part — so
  // a degenerate (empty-body or single-row) table is proven to survive a load→save round-trip.
  roundtripSpecTableFacts(spec) {
    const tableFacts = parts =>
      Object.keys(parts)
        .filter(n => /^xl\/tables\/table\d+\.xml$/.test(n))
        .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))
        .map(n => {
          const xml = parts[n];
          return {
            ref: (xml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
            wellFormed: !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml),
          };
        });
    const write = tableFacts(partMapOf(writeXlsx(buildFrom(spec))));
    let loadOk = true;
    let loadError = null;
    let roundtrip = [];
    try {
      const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
      roundtrip = tableFacts(partMapOf(writeXlsx(reloaded)));
    } catch (e) {
      loadOk = false;
      loadError = String((e && e.message) || e);
    }
    return {write, roundtrip, loadOk, loadError};
  },

  // Author a five-column table, round-trip it, and report the loaded column count and names — the
  // reader must expose every column in order, not truncate to a fixed cap.
  wideTableColumnReadReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'Wide',
      ref: 'A1',
      columns: [{name: 'C1'}, {name: 'C2'}, {name: 'C3'}, {name: 'C4'}, {name: 'C5'}],
      rowCount: 2,
    });
    const table = readXlsx(writeXlsx(wb)).getWorksheet('S').tables[0];
    return {colCount: table.columns.length, colNames: table.columns.map(c => c.name)};
  },

  // Add a table and a list validation to each of five sheets, then report that the package writes
  // with unique table part ids, reloads with every table present, and keeps the first sheet's
  // validation — the loop over many sheets must not collide table ids or strip validations.
  multiSheetTableReport() {
    const wb = new Workbook();
    for (let i = 1; i <= 5; i++) {
      const s = wb.addWorksheet(`Sheet${i}`);
      s.addTable({name: `Tbl${i}`, ref: 'A1', columns: [{name: 'Col'}], rowCount: 1});
      s.addDataValidation('C1', {type: 'list', allowBlank: true, formulae: ['"a,b,c"']});
    }
    let writeOk = true;
    let writeError = null;
    let buffer = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e && e.message) || e);
    }
    if (!writeOk) return {writeOk, writeError, reloadOk: false, tableCount: null, idsUnique: false, firstSheetDvSurvives: false};

    const parts = partMapOf(buffer);
    const ids = Object.keys(parts)
      .filter(n => /^xl\/tables\/table\d+\.xml$/.test(n))
      .map(n => (parts[n].match(/<table\b[^>]*\bid="([^"]*)"/) || [])[1]);
    const idsUnique = ids.length > 0 && ids.every(x => x != null) && new Set(ids).size === ids.length;

    let reloadOk = true;
    let tableCount = 0;
    let firstSheetDvSurvives = false;
    try {
      const back = readXlsx(buffer);
      tableCount = back.worksheets.reduce((n, s) => n + s.tables.length, 0);
      firstSheetDvSurvives = !!back.getWorksheet('Sheet1')?.dataValidationAt('C1');
    } catch {
      reloadOk = false;
    }
    return {writeOk, writeError, reloadOk, tableCount, idsUnique, firstSheetDvSurvives};
  },

  // Write three tables — one with a real built-in style name, one with no style, one with the
  // sentinel theme "None" — and report the tableStyleInfo name attribute (or null when absent) each
  // emits, plus whether the "None" table kept its showRowStripes flag. Theme "None" must mean an
  // unstyled table (no name attribute), not a bogus name="None" referencing a non-existent style.
  tableStyleThemeReport() {
    const styleInfoOf = (style) => {
      const wb = new Workbook();
      wb.addWorksheet('S').addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1, style});
      let ok = true;
      let tag = null;
      try {
        const parts = partMapOf(writeXlsx(wb));
        const part = Object.keys(parts).find(n => /^xl\/tables\/table\d+\.xml$/.test(n));
        tag = (parts[part].match(/<tableStyleInfo[^>]*\/?>/) || [])[0] ?? null;
      } catch (e) {
        ok = false;
        tag = String((e && e.message) || e);
      }
      const name = tag && ok ? ((tag.match(/\bname="([^"]*)"/) || [])[1] ?? null) : null;
      const hasStripes = !!(tag && ok && /\bshowRowStripes="1"/.test(tag));
      return {ok, name, hasStripes, tag};
    };
    return {
      real: styleInfoOf({name: 'TableStyleMedium2'}),
      // An explicit style object with no name is OOXML's "unstyled" — distinct from omitting the
      // style entirely, which a freshly-authored table fills with Excel's default (TableStyleMedium2).
      nullTheme: styleInfoOf({}),
      none: styleInfoOf({name: 'None', showRowStripes: true}),
    };
  },

  // Author a table from a header-name list (which may contain collisions), write it, and report the
  // column names emitted into the table part plus whether they are unique. OOXML requires unique
  // tableColumn names; colliding inputs must be disambiguated deterministically, not written verbatim
  // into a corrupt file → { ok, writtenNames, uniqueNames }.
  tableDuplicateColumnNamesReport(headers) {
    const wb = new Workbook();
    let ok = true;
    let writtenNames = null;
    try {
      wb.addWorksheet('S').addTable({
        name: 'T',
        ref: 'A1',
        columns: headers.map(name => ({name})),
        rowCount: 1,
      });
      const parts = partMapOf(writeXlsx(wb));
      const part = Object.keys(parts).find(n => /^xl\/tables\/table\d+\.xml$/.test(n));
      writtenNames = [...parts[part].matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map(m => m[1]);
    } catch (e) {
      ok = false;
      writtenNames = String((e && e.message) || e);
    }
    const uniqueNames =
      Array.isArray(writtenNames) &&
      new Set(writtenNames.map(n => n.toLowerCase())).size === writtenNames.length;
    return {ok, writtenNames, uniqueNames};
  },

  // Define four adjacent columns with identical width and outline level, write, and report whether
  // the write and reload succeed and how many <col> spans the part carries. Equivalent adjacent
  // columns must coalesce into fewer <col> spans than columns, without the collapse pass throwing.
  equivalentColumnCollapseReport() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    for (let i = 1; i <= 4; i++) {
      const col = s.getColumn(i);
      col.width = 12;
      col.outlineLevel = 1;
    }
    let writeOk = true;
    let writeError = null;
    let buffer = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e && e.message) || e);
    }
    if (!writeOk) return {writeOk, writeError, reloadOk: false, colSpanCount: null};

    const parts = partMapOf(buffer);
    const sheetPart = Object.keys(parts).find(n => /xl\/worksheets\/sheet\d+\.xml$/.test(n));
    const colsBlock = (parts[sheetPart].match(/<cols>[\s\S]*?<\/cols>/) || [])[0] ?? '';
    const colSpanCount = (colsBlock.match(/<col\b/g) || []).length;

    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {writeOk, writeError, reloadOk, colSpanCount};
  },

  // Author a two-column table whose first column carries a numFmt style, append two data rows, then
  // round-trip and report the numFmt read back on each column's body cells → { writeOk, reloadOk,
  // writeError, styledBody, unstyledBody }. The per-column style must bake into the styled column's
  // body cells and leave the unstyled column untouched.
  tableColumnStyleReport(numFmt) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    const table = s.addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'Amount', style: {numFmt}}, {name: 'Label'}],
      rowCount: 0,
    });
    s.getCell('A1').value = 'Amount';
    s.getCell('B1').value = 'Label';
    table.addRow([1234.5, 'x']);
    table.addRow([6789, 'y']);

    let writeOk = true;
    let writeError = null;
    let buffer = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e && e.message) || e);
    }
    if (!writeOk) return {writeOk, writeError, reloadOk: false, styledBody: null, unstyledBody: null};

    let reloadOk = true;
    let styledBody = null;
    let unstyledBody = null;
    try {
      const back = readXlsx(buffer).getWorksheet('S');
      styledBody = [back.getCell('A2').numFmt ?? null, back.getCell('A3').numFmt ?? null];
      unstyledBody = [back.getCell('B2').numFmt ?? null, back.getCell('B3').numFmt ?? null];
    } catch (e) {
      reloadOk = false;
      writeError = String((e && e.message) || e);
    }
    return {writeOk, writeError, reloadOk, styledBody, unstyledBody};
  },

  // Build a table-bearing spec, round-trip it through a write→read, fetch the named table on the
  // reloaded model, append the requested rows, then re-write and re-read to report the final row
  // count. A table read from a file must expose its data rows and accept appends exactly like a
  // freshly-created one → { hasTable, loadedRowCount, addError, committed, finalRowCount }.
  roundtripTableAppend(spec, {tableName, appendRows}) {
    const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
    let table = null;
    for (const s of reloaded.worksheets) {
      const found = s.getTable(tableName);
      if (found) { table = found; break; }
    }
    const hasTable = table !== null;
    if (!hasTable) {
      return {hasTable, loadedRowCount: null, addError: null, committed: false, finalRowCount: null};
    }
    const loadedRowCount = table.rowCount;

    let addError = null;
    for (const row of appendRows) {
      try {
        table.addRow(row);
      } catch (e) {
        addError = String((e && e.message) || e);
        break;
      }
    }

    let committed = false;
    let finalRowCount = null;
    if (addError === null) {
      try {
        const out = writeXlsx(reloaded);
        committed = true;
        const back = readXlsx(out);
        for (const s of back.worksheets) {
          const found = s.getTable(tableName);
          if (found) { finalRowCount = found.rowCount; break; }
        }
      } catch (e) {
        addError = String((e && e.message) || e);
      }
    }
    return {hasTable, loadedRowCount, addError, committed, finalRowCount};
  },

  // Author a table over A1:B3, populate its cells, load the package, edit a body cell (B2 → 999),
  // and re-write — reporting that both writes and the reload succeed, the table part and its unique
  // worksheet relationship survive, and the edited value reads back. Editing a cell inside a table's
  // range must not truncate or corrupt the table part or its rels.
  tableCellEditRoundtrip() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.addTable({name: 'T', ref: 'A1', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
    s.getCell('A1').value = 'H1';
    s.getCell('B1').value = 'H2';
    s.getCell('A2').value = 'a';
    s.getCell('B2').value = 1;
    s.getCell('A3').value = 'b';
    s.getCell('B3').value = 2;

    let writeOk = true;
    let writeError = null;
    let firstBuffer = null;
    try {
      firstBuffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e && e.message) || e);
    }
    if (!writeOk) {
      return {writeOk, writeError, reloadOk: false, hasTablePart: false, tablePresent: false, editedValue: null, relUnique: false};
    }

    let reloadOk = true;
    let hasTablePart = false;
    let tablePresent = false;
    let editedValue = null;
    let relUnique = false;
    try {
      const reloaded = readXlsx(firstBuffer);
      const sheet = reloaded.getWorksheet('S');
      sheet.getCell('B2').value = 999;
      const out = writeXlsx(reloaded);
      const parts = partMapOf(out);
      const tablePart = Object.keys(parts).find(n => /^xl\/tables\/table\d+\.xml$/.test(n));
      hasTablePart = tablePart !== undefined;
      const relPart = Object.keys(parts).find(n => /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n));
      const relIds = relPart ? [...parts[relPart].matchAll(/Id="([^"]*)"/g)].map(m => m[1]) : [];
      relUnique = relIds.length > 0 && new Set(relIds).size === relIds.length;
      const back = readXlsx(out);
      const backSheet = back.getWorksheet('S');
      tablePresent = backSheet.tables.some(t => t.name === 'T');
      editedValue = backSheet.getCell('B2').value;
    } catch (e) {
      reloadOk = false;
      writeError = String((e && e.message) || e);
    }
    return {writeOk, writeError, reloadOk, hasTablePart, tablePresent, editedValue, relUnique};
  },

  // Write a table whose first column name embeds CR/LF line breaks, then report the first
  // <tableColumn> tag and whether it carries a raw (unescaped) CR or LF. A raw control char in an
  // attribute value is not preserved by XML normalisation (a CR reparses as a space) and makes the
  // package suspect, so the name must be emitted XML-escaped (&#13;&#10;), not raw.
  tableColumnNameControlChars() {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'Test\r\nmultiple\r\nlines'}, {name: 'Plain'}],
      rowCount: 1,
    });
    let writeOk = true;
    let writeError = null;
    let firstColumnTag = null;
    let rawControlChars = null;
    try {
      const parts = partMapOf(writeXlsx(wb));
      const part = Object.keys(parts).find(n => /^xl\/tables\/table\d+\.xml$/.test(n));
      firstColumnTag = (parts[part].match(/<tableColumn\b[^>]*\/?>/) || [])[0] ?? null;
      rawControlChars = firstColumnTag === null ? null : /[\r\n]/.test(firstColumnTag);
    } catch (e) {
      writeOk = false;
      writeError = String((e && e.message) || e);
    }
    return {writeOk, writeError, firstColumnTag, rawControlChars};
  },

  // Read every table part of a fixture, then load→save it and read the re-emitted parts, reporting
  // each table's autoFilter / header-row / totals-row / column-count facts before and after. A no-op
  // round-trip of a table that has no autoFilter must not inject one, flip the header row off, or turn
  // totalsRowShown on; a table that does have one must keep its ref and column count.
  roundtripFixtureTableXml(rel) {
    const facts = xml => ({
      hasAutoFilter: /<(?:\w+:)?autoFilter\b/.test(xml),
      autoFilterRef: (xml.match(/<(?:\w+:)?autoFilter\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
      headerRowCount: (xml.match(/\bheaderRowCount="([^"]*)"/) || [])[1] ?? null,
      totalsRowShown: (xml.match(/\btotalsRowShown="([^"]*)"/) || [])[1] ?? null,
      columnCount: (xml.match(/<tableColumns\b[^>]*\bcount="([^"]*)"/) || [])[1] ?? null,
    });
    const tablePartsInOrder = parts =>
      Object.keys(parts)
        .filter(n => /^xl\/tables\/table\d+\.xml$/.test(n))
        .sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))
        .map(n => parts[n]);
    const buffer = fs.readFileSync(path.join(FIXTURES_ROOT, rel));
    const source = tablePartsInOrder(partMapOf(buffer));
    const rewritten = tablePartsInOrder(partMapOf(writeXlsx(readXlsx(buffer))));
    return {
      tables: source.map((xml, i) => ({
        name: (xml.match(/<table\b[^>]*\bname="([^"]*)"/) || [])[1] ?? null,
        source: facts(xml),
        rewritten: rewritten[i] ? facts(rewritten[i]) : null,
      })),
    };
  },

  // Author a table whose display name differs from its internal name, then report the displayName
  // written into the table part and the internal/display names read back from the reloaded model —
  // a serializer that mis-keys the property drops the display name to the internal default.
  tableDisplayNameReport(display) {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'MyTable',
      displayName: display,
      ref: 'A1',
      columns: [{name: 'C'}],
      rowCount: 1,
    });
    const buffer = writeXlsx(wb);
    const part = partMapOf(buffer)['xl/tables/table1.xml'] || '';
    const writtenDisplayName = (part.match(/\bdisplayName="([^"]*)"/) || [])[1] ?? null;
    const table = readXlsx(buffer).getWorksheet('S').tables[0];
    return {
      writtenDisplayName,
      reloadedDisplayName: table ? table.displayName : null,
      reloadedName: table ? table.name : null,
    };
  },

  // Build a formula-bearing spec, write it, read it back, and report each cell as
  // { formula, sharedFormula, result } — mirroring the oracle. A shared-formula clone reads back a
  // concrete formula (the master's, translated to the clone's address) while retaining its master
  // reference under `sharedFormula`; a plain formula master carries no `sharedFormula`.
  roundtripFormulas(spec) {
    const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
    const out = {};
    for (const s of spec.sheets || []) {
      const sheet = reloaded.getWorksheet(s.name);
      for (const c of s.cells || []) {
        const v = sheet ? sheet.getCell(c.ref).value : null;
        const obj = v && typeof v === 'object';
        out[c.ref] = {
          formula: obj && 'formula' in v ? v.formula : null,
          sharedFormula: obj && 'sharedFormula' in v ? v.sharedFormula : null,
          result: obj && 'result' in v ? v.result ?? null : null,
        };
      }
    }
    return out;
  },

  // Build a shared-formula sheet (master B1 filled down to B2/B3), then report two things: whether a
  // read → write round-trip preserves the dependents as formula cells, and whether splicing a column
  // into the loaded sheet writes without throwing. The clone's master reference is an address the
  // rewrite does not yet re-anchor on a structural edit, so the splice is the known-open here.
  sharedFormulaRoundtripAndSplice() {
    const build = () => {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      sheet.getCell('A1').value = 1;
      sheet.getCell('A2').value = 2;
      sheet.getCell('A3').value = 3;
      sheet.getCell('B1').value = {formula: 'A1*2', result: 2};
      sheet.getCell('B2').value = {sharedFormula: 'B1', result: 4};
      sheet.getCell('B3').value = {sharedFormula: 'B1', result: 6};
      return wb;
    };
    const buffer = writeXlsx(build());

    let roundtripError = null;
    let preservedFormulas = null;
    try {
      const reread = readXlsx(buffer);
      writeXlsx(reread);
      const s = reread.getWorksheet('S');
      preservedFormulas = ['B2', 'B3'].every(ref => {
        const v = s.getCell(ref).value;
        return !!(v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v));
      });
    } catch (e) {
      roundtripError = String((e && e.message) || e);
    }

    let spliceError = null;
    try {
      const reread = readXlsx(buffer);
      reread.getWorksheet('S').spliceColumns(1, 0, []);
      writeXlsx(reread);
    } catch (e) {
      spliceError = String((e && e.message) || e);
    }

    return {
      roundtripOk: roundtripError === null,
      roundtripError,
      preservedFormulas,
      spliceOk: spliceError === null,
      spliceError,
    };
  },

  tryWriteWorkbook(spec) {
    let workbook;
    try {
      workbook = buildFrom(spec);
    } catch (error) {
      if (error.notImplemented) throw error;
      return {ok: false, phase: 'build', error: String((error && error.message) || error)};
    }
    let buffer;
    try {
      buffer = writeXlsx(workbook);
    } catch (error) {
      if (error.notImplemented) throw error;
      return {ok: false, phase: 'write', error: String((error && error.message) || error)};
    }
    // Report which cells survived the round-trip, so a case can prove a bad cell (e.g. an Invalid
    // Date, written value-less) did not drop its siblings.
    const reread = readXlsx(buffer);
    const survivingCells = {};
    for (const s of spec.sheets || []) {
      const sheet = reread.getWorksheet(s.name);
      survivingCells[s.name] = (s.cells || [])
        .filter(c => sheet && sheet.getCell(c.ref).value !== null)
        .map(c => c.ref);
    }
    return {ok: true, byteLength: buffer.byteLength ?? buffer.length, survivingCells};
  },

  // Author per-cell protection (and optionally protect the sheet), write, then read back →
  // { readBack, hasApplyProtection, sheetProtection, sheetProtectionAttrs }. Reports whether an
  // explicitly-unlocked cell round-trips as locked=false, whether the style record carries the
  // flag (applyProtection + <protection> in cellXfs), and the emitted <sheetProtection> that
  // makes the locked flags enforceable.
  authorCellProtection(cells = [], protect = null, {rows = [], columns = []} = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const c of cells) {
      const cell = sheet.getCell(c.ref);
      cell.value = c.value ?? c.ref;
      if (c.protection !== undefined) cell.protection = c.protection;
    }
    // Whole-column / whole-row protection: the model carries protection per cell, so realize an
    // unlocked band by stamping its flag onto each listed cell that falls in the band — the same
    // end-state a per-cell override yields (column-scope inheritance is a separate capability).
    // Applied after the per-cell settings so the band-level flag is what the case asserts.
    for (const col of columns) {
      for (const c of cells) if (decodeAddress(c.ref).col === col.index) sheet.getCell(c.ref).protection = col.protection;
    }
    for (const r of rows) {
      for (const c of cells) if (decodeAddress(c.ref).row === r.index) sheet.getCell(c.ref).protection = r.protection;
    }
    if (protect) sheet.protect(protect.password ?? undefined, protect.options ?? {});
    const buffer = writeXlsx(workbook);
    const parts = partMapOf(buffer);
    const styles = parts['xl/styles.xml'] || '';
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const sheetProtection = (sheetXml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0] || null;

    const reread = readXlsx(buffer);
    const sheet2 = reread.getWorksheet('S');
    const readBack = {};
    for (const c of cells) {
      const p = sheet2.getCell(c.ref).protection;
      readBack[c.ref] = p ? {locked: p.locked ?? null} : null;
    }
    return {
      readBack,
      hasApplyProtection: /applyProtection="1"/.test(styles) && /<protection\b/.test(styles),
      sheetProtection,
      sheetProtectionAttrs: sheetProtection ? attrsOf(sheetProtection) : null,
    };
  },

  // Password-protect a worksheet twice under Node and report the emitted protection facts →
  // { threw, algorithm, hasHash, hasSalt, spinCount, selectLockedCells, selectUnlockedCells,
  // saltsDiffer }. Proves protect succeeds without a browser-random error, emits a well-formed
  // password credential, honors the requested options, and salts with real randomness (two
  // protects with the same password differ).
  worksheetPasswordProtectionReport(password = 'secret') {
    const protectOnce = () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.protect(password, {selectLockedCells: false, selectUnlockedCells: false});
      const xml = partMapOf(writeXlsx(wb))['xl/worksheets/sheet1.xml'] || '';
      return (xml.match(/<sheetProtection\b[^>]*\/>/) || [''])[0];
    };
    let first = '';
    let second = '';
    try {
      first = protectOnce();
      second = protectOnce();
    } catch (e) {
      return {
        threw: String((e && e.message) || e),
        algorithm: null, hasHash: false, hasSalt: false, spinCount: null,
        selectLockedCells: null, selectUnlockedCells: null, saltsDiffer: false,
      };
    }
    const a = attrsOf(first);
    const b = attrsOf(second);
    return {
      threw: null,
      algorithm: a.algorithmName ?? null,
      hasHash: !!a.hashValue,
      hasSalt: !!a.saltValue,
      spinCount: a.spinCount ?? null,
      selectLockedCells: a.selectLockedCells ?? null,
      selectUnlockedCells: a.selectUnlockedCells ?? null,
      saltsDiffer: !!a.saltValue && !!b.saltValue && a.saltValue !== b.saltValue,
    };
  },

  // Protect a worksheet, write it, read it back, then write the reloaded workbook again, reporting
  // the <sheetProtection> attributes from BOTH writes → { first, second }. Proves the reader carries
  // sheet-level protection back into the model rather than silently dropping it on a passthrough
  // save — the second write must still emit protection, with the agile credential preserved verbatim
  // (no plaintext password survives to re-hash) and the permissive flags intact.
  sheetProtectionRoundtrip(
    password = 'secret',
    options = {sort: true, autoFilter: true, selectLockedCells: false}
  ) {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.protect(password ?? undefined, options ?? {});

    const buf1 = writeXlsx(wb);
    const buf2 = writeXlsx(readXlsx(buf1));
    const protAttrs = buf => {
      const xml = partMapOf(buf)['xl/worksheets/sheet1.xml'] || '';
      const el = (xml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0];
      return el ? attrsOf(el) : null;
    };
    return {first: protAttrs(buf1), second: protAttrs(buf2)};
  },

  // Copy-on-write style aliasing family. Each cell owns its facet fields and every setter REPLACES
  // the field (the readonly facet types forbid in-place mutation of a shared record), so mutating
  // one cell's facet — even a cell that shared a style with siblings on disk — cannot bleed onto a
  // sibling. These methods prove that end-to-end through the real write→read path.

  // Assign one base style object to two cells, then spread-reassign one cell's font color →
  // { a1Color, a2Color, bled }. The sibling given the same base must keep its original font.
  sharedBaseStyleFontMutation() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const base = {font: {name: 'Arial', size: 11}};
    sheet.getCell('A1').value = 'YES';
    sheet.getCell('A2').value = 'NO';
    applyStyle(sheet.getCell('A1'), base);
    applyStyle(sheet.getCell('A2'), base);
    sheet.getCell('A1').font = {...sheet.getCell('A1').font, color: {argb: 'FF00FF00'}};
    const s = readXlsx(writeXlsx(workbook)).getWorksheet('S');
    const colorOf = ref => {
      const f = s.getCell(ref).font;
      return f && f.color ? f.color.argb ?? null : null;
    };
    const a1Color = colorOf('A1');
    const a2Color = colorOf('A2');
    return {a1Color, a2Color, bled: a2Color === 'FF00FF00'};
  },

  // Author three cells sharing one style record, load, border ONE, round-trip →
  // { a1, a2, a3, bled }. Only the targeted cell may gain a border.
  loadMutateCellBorder() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const r of [1, 2, 3]) {
      const c = sheet.getCell(`A${r}`);
      c.value = 'x';
      c.font = {bold: true};
    }
    const loaded = readXlsx(writeXlsx(workbook));
    loaded.getWorksheet('S').getCell('A1').border = {
      top: {style: 'thin'}, left: {style: 'thin'}, bottom: {style: 'thin'}, right: {style: 'thin'},
    };
    const s = readXlsx(writeXlsx(loaded)).getWorksheet('S');
    const hasBorder = ref => {
      const b = s.getCell(ref).border;
      return !!(b && b.top && b.top.style);
    };
    return {a1: hasBorder('A1'), a2: hasBorder('A2'), a3: hasBorder('A3'), bled: hasBorder('A2') || hasBorder('A3')};
  },

  // Author two cells with one shared fill, load, replace ONE cell's fill, read the sibling in
  // memory and after write-back → { sibling, mutatedTo, original, bled, diskSibling, diskBled }.
  loadMutateCellStyle({sharedFill = 'FFFF0000', mutateTo = 'FF00FF00'} = {}) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: sharedFill}};
    s.getCell('A1').fill = fill;
    s.getCell('B1').fill = fill; // identical formatting → one shared style index on disk
    const fgOf = cell => (cell.fill && cell.fill.fgColor ? cell.fill.fgColor.argb ?? null : null);

    const wb2 = readXlsx(writeXlsx(wb));
    const s2 = wb2.getWorksheet('S');
    s2.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: mutateTo}};
    const sibling = fgOf(s2.getCell('B1'));

    const diskSibling = fgOf(readXlsx(writeXlsx(wb2)).getWorksheet('S').getCell('B1'));
    return {
      sibling, mutatedTo: mutateTo, original: sharedFill,
      bled: sibling === mutateTo, diskSibling, diskBled: diskSibling === mutateTo,
    };
  },

  // Author two cells with one shared font, load, spread-reassign ONE cell's font color, read the
  // sibling → { edited, sibling, original, mutatedTo, bled }.
  loadMutateCellFont({original = 'FF000000', mutateTo = 'FFFF0000'} = {}) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    const font = {name: 'Arial', size: 12, color: {argb: original}};
    s.getCell('A1').value = 'a';
    s.getCell('A1').font = font;
    s.getCell('B1').value = 'b';
    s.getCell('B1').font = font; // identical formatting → one shared style index on disk

    const s2 = readXlsx(writeXlsx(wb)).getWorksheet('S');
    s2.getCell('A1').font = {...s2.getCell('A1').font, color: {argb: mutateTo}};
    const colorOf = cell => (cell.font && cell.font.color ? cell.font.color.argb ?? null : null);
    const sibling = colorOf(s2.getCell('B1'));
    return {edited: colorOf(s2.getCell('A1')), sibling, original, mutatedTo: mutateTo, bled: sibling === mutateTo};
  },

  // Load two cells sharing one style record, set ONE style facet (alignment | numFmt | protection)
  // on one via its setter, and report whether it bled into the sibling in memory and on disk →
  // { facet, target, sibling, original, bled, diskSibling, diskBled }. The remaining facets of the
  // copy-on-write family, alongside fill/font/border above.
  loadMutateCellFacet(facet = 'alignment') {
    const readFacet = {
      alignment: c => (c.alignment && c.alignment.horizontal) || null,
      numFmt: c => c.numFmt || null,
      protection: c => (c.protection && typeof c.protection.locked === 'boolean' ? c.protection.locked : null),
    }[facet];
    const apply = {
      alignment: c => { c.alignment = {horizontal: 'center'}; },
      numFmt: c => { c.numFmt = '#,##0'; },
      protection: c => { c.protection = {locked: false}; },
    }[facet];
    if (!readFacet) throw new Error(`unknown style facet: ${facet}`);

    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    const base = {numFmt: '0.00'}; // one identical non-default style → both cells dedup to one xf index
    applyStyle(s.getCell('A1'), base);
    applyStyle(s.getCell('B1'), base);

    const wb2 = readXlsx(writeXlsx(wb));
    const s2 = wb2.getWorksheet('S');
    const original = readFacet(s2.getCell('B1'));
    apply(s2.getCell('A1'));
    const target = readFacet(s2.getCell('A1'));
    const sibling = readFacet(s2.getCell('B1'));

    const diskSibling = readFacet(readXlsx(writeXlsx(wb2)).getWorksheet('S').getCell('B1'));
    return {facet, target, sibling, original, bled: sibling !== original, diskSibling, diskBled: diskSibling !== original};
  },

  // --- CSV (src/io/csv) -------------------------------------------------------------------------
  // The contract mirrors the oracle's ExcelJS-shaped options; here they translate onto the rewrite's
  // cleaner CsvReadOptions/CsvWriteOptions. A read yields a JSON-serializable 2-D array of typed
  // values (Date → { date: iso }, error → { error }, else the scalar or null), matching the oracle so
  // the same cases assert unchanged.

  csvRead({csv, options} = {}) {
    try {
      const wb = readCsv(csv, translateCsvReadOptions(options));
      const sheet = wb.worksheets[0];
      const rows = [];
      if (sheet) {
        for (const {cells} of sheet.rows()) {
          let width = 0;
          for (const cell of cells) if (cell.col > width) width = cell.col;
          const fields = new Array(width).fill(null);
          for (const cell of cells) fields[cell.col - 1] = normalizeCsvValue(cell.value);
          rows.push(fields);
        }
      }
      return {ok: true, error: null, rows};
    } catch (e) {
      return {ok: false, error: String((e && e.message) || e), rows: []};
    }
  },

  csvWrite({spec = {}, options} = {}) {
    try {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      for (const row of spec.rows || []) sheet.addRow((row || []).map(specCsvValue));
      const text = writeCsvText(wb, translateCsvWriteOptions(options));
      return {ok: true, error: null, text};
    } catch (e) {
      return {ok: false, error: String((e && e.message) || e), text: null};
    }
  },

  csvWriteSheetSelection(sheetName) {
    const wb = new Workbook();
    wb.addWorksheet('First').addRow(['a', 1]);
    const second = wb.addWorksheet('Second');
    second.addRow(['b', 2]);
    second.addRow(['c', 3]);
    let error = null;
    let text = null;
    try {
      text = writeCsvText(wb, sheetName === undefined ? {} : {sheetName});
    } catch (e) {
      error = String((e && e.message) || e);
    }
    return {ok: error === null, error, text, rowCount: text ? text.split(/\r?\n/).filter(Boolean).length : 0};
  },

  csvReadMapReport() {
    const csv = 'id,amount\n007,32.5\n008,40';
    const read = map => {
      const wb = readCsv(csv, map ? {map} : {});
      const sheet = wb.worksheets[0];
      const a = sheet ? sheet.getCell('A2').value : null;
      const b = sheet ? sheet.getCell('B2').value : null;
      return {a, aType: typeof a, b, bType: typeof b};
    };
    return {default: read(null), identity: read(v => v)};
  },

  csvWriteEncodingReport({encoding = 'utf16le', text = 'café'} = {}) {
    const EMOJI = '😀🎉';
    const CJK = '日本語テスト';

    const fidelityWb = new Workbook();
    fidelityWb.addWorksheet('S').addRow([EMOJI, CJK]);
    const reread = readCsv(writeCsv(fidelityWb)).worksheets[0];
    const emojiRoundtrips =
      !!reread && reread.getCell('A1').value === EMOJI && reread.getCell('B1').value === CJK;

    const encodedWb = new Workbook();
    encodedWb.addWorksheet('S').addRow([text]);
    const encodedBuffer = Buffer.from(writeCsv(encodedWb, {encoding}));
    const decodesAsRequested = encodedBuffer.toString(encoding).replace(/\r?\n$/, '') === text;
    const decodesAsUtf8 = encodedBuffer.toString('utf8').replace(/\r?\n$/, '') === text;

    return {emojiRoundtrips, requestedEncoding: encoding, decodesAsRequested, decodesAsUtf8};
  },

  csvNonAsciiEncodingReport(text = 'שלום') {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = text;
    sheet.getCell('B1').value = 'world';
    const buffer = Buffer.from(writeCsv(wb));
    const hasBom = buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
    const body = hasBom ? buffer.subarray(3) : buffer;
    return {hasBom, bytesDecodeToText: body.toString('utf8').startsWith(text)};
  },
};

// A JSON-serializable view of a read-back CSV cell value, mirroring the oracle's normalizeCsvValue.
const normalizeCsvValue = v => {
  if (v instanceof Date) return {date: Number.isNaN(v.getTime()) ? null : v.toISOString()};
  if (v && typeof v === 'object' && 'error' in v) return {error: v.error};
  return v ?? null;
};

// A declarative CSV write-spec cell → a live model value: { date } → Date, { formula, result } →
// formula value, { error } → error value, primitive passes through.
const specCsvValue = c => {
  if (c && typeof c === 'object') {
    if (c.date) return new Date(c.date);
    if ('formula' in c) return {formula: c.formula, result: c.result};
    if ('error' in c) return {error: c.error};
  }
  return c;
};

// The oracle's ExcelJS-shaped read options → the rewrite's CsvReadOptions.
const translateCsvReadOptions = (options = {}) => {
  const parser = options.parserOptions || {};
  const translated = {};
  if (parser.delimiter !== undefined) translated.delimiter = parser.delimiter;
  if (parser.headers) translated.headers = true;
  if (typeof options.map === 'function') translated.map = options.map;
  return translated;
};

// The oracle's ExcelJS-shaped write options → the rewrite's CsvWriteOptions.
const translateCsvWriteOptions = (options = {}) => {
  const formatter = options.formatterOptions || {};
  const translated = {};
  if (formatter.delimiter !== undefined) translated.delimiter = formatter.delimiter;
  if (options.dateFormat !== undefined) translated.dateFormat = options.dateFormat;
  if (options.dateUTC !== undefined) translated.dateUTC = options.dateUTC;
  return translated;
};

export default new Proxy(impl, {
  get(target, prop, receiver) {
    if (prop in target || typeof prop === 'symbol') {
      return Reflect.get(target, prop, receiver);
    }
    return () => {
      throw notImplemented(`capability "${prop}" is not implemented yet`);
    };
  },
});
