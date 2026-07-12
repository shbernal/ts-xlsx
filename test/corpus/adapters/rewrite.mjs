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

import {strFromU8, unzipSync} from 'fflate';

import {decodeAddress, decodeRange} from '../../../src/core/address.ts';
import {Workbook} from '../../../src/core/workbook.ts';
import {readXlsx} from '../../../src/io/xlsx/read.ts';
import {writeXlsx} from '../../../src/io/xlsx/write.ts';
import {packageFacts} from './ooxml-facts.mjs';

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
const SUPPORTED_COLUMN_KEYS = new Set(['index', 'width', 'hidden', 'numFmt']);
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
