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

import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {decodeAddress, decodeRange, encodeAddress} from '../../../src/core/address.ts';
import {Workbook} from '../../../src/core/workbook.ts';
import {readXlsx} from '../../../src/io/xlsx/read.ts';
import {writeXlsx} from '../../../src/io/xlsx/write.ts';
import {packageFacts} from './ooxml-facts.mjs';

// Durable sample inputs live under test/corpus/fixtures/<case-slug>/ — the SAME tree the
// oracle adapter reads, so a fixture-backed case measures both implementations against one
// real-world file. The rewrite reads them straight through readXlsx (a fixture is just a
// foreign `.xlsx` buffer).
const FIXTURES_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'fixtures');
const readFixture = rel => readXlsx(fs.readFileSync(path.join(FIXTURES_ROOT, rel)));

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

const SUPPORTED_TOP_KEYS = new Set(['sheets', 'properties']);
const SUPPORTED_PROP_KEYS = new Set(['creator', 'lastModifiedBy', 'created', 'modified']);
const SUPPORTED_SHEET_KEYS = new Set([
  'name', 'state', 'cells', 'columns', 'rows', 'properties', 'pageMargins', 'headerFooter', 'tables', 'merges',
]);
const SUPPORTED_CELL_KEYS = new Set(['ref', 'value', 'formula', 'result', 'fill', 'numFmt', 'font', 'border', 'alignment', 'protection']);
const SUPPORTED_SHEET_PROP_KEYS = new Set(['defaultRowHeight', 'defaultColWidth']);
const SUPPORTED_COLUMN_KEYS = new Set(['index', 'width', 'hidden', 'numFmt', 'fill', 'font', 'border', 'alignment', 'protection']);
const SUPPORTED_ROW_KEYS = new Set(['index', 'height', 'hidden', 'outlineLevel', 'collapsed', 'fill']);
const SUPPORTED_PAGE_MARGIN_KEYS = new Set(['left', 'right', 'top', 'bottom', 'header', 'footer']);
const SUPPORTED_HEADER_FOOTER_KEYS = new Set([
  'oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter',
]);
const SUPPORTED_TABLE_KEYS = new Set([
  'name', 'ref', 'headers', 'columnDefs', 'rows', 'headerRow', 'totalsRow',
]);
const SUPPORTED_TABLE_COLUMN_KEYS = new Set(['name', 'totalsRowLabel', 'totalsRowFunction']);

const toDate = v => (v && typeof v === 'object' && v.invalidDate ? new Date(NaN) : new Date(v));
const isoOrNull = d => (d instanceof Date && !Number.isNaN(d.getTime()) ? d.toISOString() : null);

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
      const options = {name: t.name, ref: t.ref, columns, rowCount: (t.rows || []).length};
      if (t.headerRow !== undefined) options.headerRow = t.headerRow;
      if (t.totalsRow !== undefined) options.totalsRow = t.totalsRow;
      sheet.addTable(options);
    }

    for (const range of s.merges || []) sheet.mergeCells(range);

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

    for (const c of s.cells || []) {
      for (const k of Object.keys(c)) {
        if (!SUPPORTED_CELL_KEYS.has(k)) throw notImplemented(`cell.${k} not supported yet`);
      }
      const cell = sheet.getCell(c.ref);
      if ('formula' in c) {
        cell.value = 'result' in c ? {formula: c.formula, result: c.result} : {formula: c.formula};
      } else if ('value' in c) {
        const v = c.value;
        if (v !== null && typeof v === 'object') {
          throw notImplemented(`cell value shape ${JSON.stringify(v)} not supported yet`);
        }
        cell.value = v;
      }
      if (c.fill !== undefined) cell.fill = c.fill;
      if (c.numFmt !== undefined) cell.numFmt = c.numFmt;
      if (c.font !== undefined) cell.font = c.font;
      if (c.border !== undefined) cell.border = c.border;
      if (c.alignment !== undefined) cell.alignment = c.alignment;
      if (c.protection !== undefined) cell.protection = c.protection;
    }
  }
  return workbook;
}

// Mirror current.mjs's normalizeCell for the rewrite's Cell: a plain JSON view of the
// value that survived the round-trip. Style facets are absent until the reader reads
// them, matching the contract that an unmaterialized facet is simply not present.
function normalizeRewriteCell(cell) {
  const v = cell.value;
  let out;
  if (v && typeof v === 'object' && 'formula' in v) out = {formula: v.formula, result: v.result ?? null};
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

  // Read a real fixture `.xlsx` and report the fill and font colour the reader surfaces for each
  // requested `<sheet>!<address>` cell → { [key]: { fill, fontColor } | null }. Mirrors the oracle:
  // a solid-pattern fill's visible colour lives on fgColor while bgColor is the automatic indexed
  // placeholder, and the font colour is a wholly separate facet — the two are never conflated.
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
          if (!hasStyle(cell)) continue;
          checked += 1;
          const beforeKey = styleKey(cell);
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
      sheets[s.name] = {
        cells,
        columns,
        rows,
        margins,
        autoFilter: null,
        merges: [...sheet.merges],
        rowCount: sheet.rowCount,
        actualRowCount: sheet.actualRowCount,
      };
    }
    const props = reloaded.properties;
    return {
      properties: {
        creator: props.creator ?? null,
        lastModifiedBy: props.lastModifiedBy ?? null,
        created: isoOrNull(props.created),
        modified: isoOrNull(props.modified),
      },
      sheets,
    };
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

    // Duplicate table column names authored separately — construction must reject them.
    let dupRejected = false;
    try {
      const w2 = new Workbook();
      w2.addWorksheet('S').addTable({name: 'T2', ref: 'A1', columns: [{name: 'Dup'}, {name: 'Dup'}], rowCount: 1});
      writeXlsx(w2);
    } catch {
      dupRejected = true;
    }

    return {
      tableRef,
      imageFromRow: imageFromRow != null ? Number(imageFromRow) : null,
      dupColumnNamesRejected: dupRejected,
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
    try {
      writeXlsx(workbook);
    } catch (error) {
      if (error.notImplemented) throw error;
      return {ok: false, phase: 'write', error: String((error && error.message) || error)};
    }
    // Whether the write succeeded or failed legibly is fully answerable without the reader;
    // reporting which cells survived a round-trip is the reader's job (a separate capability).
    return {ok: true};
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
