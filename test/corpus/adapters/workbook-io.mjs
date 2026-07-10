// Workbook round-trip / package-inspection capabilities for the `current` adapter.
// This is a helper of the `current` adapter (not itself discovered by the runner,
// which only imports adapters/<name>.mjs by exact name) — so it, like the adapter,
// is allowed to know how today's `lib/` implementation is shaped. Everything it
// returns is plain JSON so cases can assert on observable behavior without touching
// implementation objects.
//
// Three keystone capabilities cover most of the corpus:
//   buildFrom(spec)        — turn a declarative workbook spec into a live workbook
//   roundtripWorkbook(spec)— build → writeBuffer → load back → normalized model
//   inspectPackage(spec)   — build → write → unzip → raw OOXML-part facts
//   tryWriteWorkbook(spec) — build → attempt write → { ok, error, … }
//
// The `spec` shape (all fields optional unless noted):
//   {
//     properties: { creator, lastModifiedBy, created, modified },   // dates: ISO strings
//     definedNames: [{ name, ranges: ["Sheet1!$A$1:$C$5", …] }],     // workbook-level names
//     sheets: [{
//       name,                                                       // required
//       cells:   [{ ref, value|formula(+result)|sharedFormula(+result)|text+hyperlink,
//                   numFmt, font, fill, alignment, note }],
//       images:  [{ range, extension? }],                           // range: "B2:D6" or {tl,br?,ext?}; extension defaults 'png'
//       columns: [{ index, width, hidden, numFmt, style }],         // index: 1-based
//       rows:    [{ index, height, hidden }],
//       pageMargins: { left, right, top, bottom, header, footer },  // any subset
//       pageSetup:   { fitToPage, fitToWidth, fitToHeight, scale, orientation, … },
//       autoFilter:  "A1:C3" | { from, to },                        // filter range
//       tables:  [{ name, ref, headers:[…], rows:[[…]], totalsRow }],
//       merges:  ["A1:B1", …],                                       // merged cell ranges
//     }],
//   }
// A cell value of { invalidDate: true } materializes `new Date(NaN)`.

import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';
import {Readable, PassThrough, Duplex} from 'node:stream';
import path from 'node:path';

const require = createRequire(import.meta.url);
const ExcelJS = require('../../../lib/exceljs.nodejs.js');
const JSZip = require('jszip');

// A 1×1 transparent PNG — the smallest valid image, so image-anchor cases can place a
// picture without shipping an image fixture. Anchor geometry is independent of pixels.
const ONE_PX_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64'
);

// Durable sample inputs live under test/corpus/fixtures/<case-slug>/. A fixture-backed
// capability takes a path relative to that root so cases never hardcode a filesystem layout.
const FIXTURES_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'fixtures');
const fixturePath = rel => path.join(FIXTURES_ROOT, rel);

const toDate = v => (v && typeof v === 'object' && v.invalidDate ? new Date(NaN) : new Date(v));

function cellValueFrom(c) {
  if ('sharedFormula' in c) return {sharedFormula: c.sharedFormula, result: c.result};
  if ('formula' in c) return {formula: c.formula, result: c.result};
  if ('hyperlink' in c) return {text: c.text ?? c.hyperlink, hyperlink: c.hyperlink};
  if (c.value && typeof c.value === 'object' && c.value.invalidDate) return new Date(NaN);
  if (c.value && typeof c.value === 'object' && c.value.date) return toDate(c.value.date);
  return c.value;
}

export function buildFrom(spec = {}) {
  const workbook = new ExcelJS.Workbook();
  const p = spec.properties || {};
  if (p.creator !== undefined) workbook.creator = p.creator;
  if (p.lastModifiedBy !== undefined) workbook.lastModifiedBy = p.lastModifiedBy;
  if (p.created !== undefined) workbook.created = toDate(p.created);
  if (p.modified !== undefined) workbook.modified = toDate(p.modified);

  for (const s of spec.sheets || []) {
    const sheet = workbook.addWorksheet(s.name);
    for (const c of s.cells || []) {
      const cell = sheet.getCell(c.ref);
      cell.value = cellValueFrom(c);
      if (c.numFmt !== undefined) cell.numFmt = c.numFmt;
      if (c.font !== undefined) cell.font = c.font;
      if (c.fill !== undefined) cell.fill = c.fill;
      if (c.alignment !== undefined) cell.alignment = c.alignment;
      if (c.note !== undefined) cell.note = c.note;
    }
    for (const col of s.columns || []) {
      const column = sheet.getColumn(col.index);
      if (col.width !== undefined) column.width = col.width;
      if (col.hidden !== undefined) column.hidden = col.hidden;
      if (col.numFmt !== undefined) column.numFmt = col.numFmt;
      if (col.style !== undefined) column.style = col.style;
    }
    for (const row of s.rows || []) {
      const r = sheet.getRow(row.index);
      if (row.height !== undefined) r.height = row.height;
      if (row.hidden !== undefined) r.hidden = row.hidden;
    }
    // Assign only the margins the spec provides — faithfully reproducing the user
    // scenario of setting a subset (do NOT pre-fill defaults; that is exactly the
    // write-side behavior under test).
    // Images: each spec image places the built-in 1×1 PNG at a range that is either a
    // string ("B2:D6") or an object anchor ({tl, br?, ext?}). Anchor geometry — not the
    // image bytes — is what these cases assert on.
    for (const img of s.images || []) {
      // `extension` defaults to 'png'; a spec may set it to a bad/missing value on purpose
      // to exercise write-side validation of the media part + content-type declaration.
      const extension = 'extension' in img ? img.extension : 'png';
      const imageId = workbook.addImage({buffer: ONE_PX_PNG, extension});
      sheet.addImage(imageId, img.range);
    }
    if (s.pageMargins) sheet.pageSetup.margins = {...s.pageMargins};
    if (s.pageSetup) Object.assign(sheet.pageSetup, s.pageSetup);
    if (s.autoFilter) sheet.autoFilter = s.autoFilter;
    if (s.headerFooter) sheet.headerFooter = {...s.headerFooter};
    for (const t of s.tables || []) {
      sheet.addTable({
        name: t.name,
        ref: t.ref,
        headerRow: t.headerRow !== false,
        totalsRow: !!t.totalsRow,
        columns: (t.headers || []).map(name => ({name, totalsRowLabel: t.totalsRow ? name : undefined})),
        rows: t.rows || [],
      });
    }
    // Merged ranges are applied after cells so a merge master keeps its assigned value.
    for (const m of s.merges || []) sheet.mergeCells(m);
  }
  // Workbook-level defined names are added after every sheet exists, since each name's
  // reference targets a sheet by name.
  for (const dn of spec.definedNames || []) {
    for (const range of dn.ranges || []) workbook.definedNames.add(range, dn.name);
  }
  return workbook;
}

// Read back the workbook's defined names as a plain { <name>: [ranges…] } map with sorted
// ranges — for asserting a name (especially a full-row/full-column span) survives a path.
function definedNamesOf(workbook) {
  const out = {};
  for (const m of workbook.definedNames.model || []) out[m.name] = [...(m.ranges || [])].sort();
  return out;
}

// Apply a sequence of structural mutations (row/column splices, row insert/duplicate) to a
// fresh single worksheet and report the observable result — for asserting that in-memory
// model edits behave predictably regardless of how many rows/columns they touch. Ops:
//   { op: 'spliceRows',    start, count, inserts?: any[][] }
//   { op: 'spliceColumns', start, count, inserts?: any[][] }
//   { op: 'insertRow',     pos, value?: any[] }
//   { op: 'duplicateRow',  start, count?, insert?: boolean }
// Returns { rowCount, columnCount, cells: { <ref>: value|null }, merges: [ranges…], error? } —
// never an implementation object. A throwing op is reported as { error } rather than propagated,
// so a case can distinguish "mutation threw" from "mutation silently did nothing".
export function mutateWorksheet({cells = [], ops = [], read = [], readStyles = []} = {}) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('S');
  for (const c of cells) {
    const cell = sheet.getCell(c.ref);
    cell.value = c.value;
    // Optional per-cell style so a case can assert that a structural edit (a row/column splice)
    // carries a cell's font/fill/numFmt to its shifted position rather than blanking it.
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
  for (const ref of read) {
    const v = sheet.getCell(ref).value;
    readCells[ref] = v ?? null;
  }
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
  // The merged-range list (as OOXML records it) after the mutations — for asserting that a row/
  // column splice SHIFTS a merged range to its new position and keeps it merged, rather than
  // leaving the range stranded at its original indices while the data moves.
  const merges = (sheet.model && sheet.model.merges) || [];
  return {rowCount: sheet.rowCount, columnCount: sheet.columnCount, cells: readCells, styles, merges, error};
}

const isoOrNull = d => (d instanceof Date && !Number.isNaN(+d) ? d.toISOString() : null);

function normalizeCell(cell) {
  const v = cell.value;
  const out = {};
  if (v && typeof v === 'object' && 'formula' in v) {
    out.formula = v.formula;
    out.result = v.result ?? null;
  } else if (v && typeof v === 'object' && 'hyperlink' in v) {
    out.hyperlink = v.hyperlink;
    out.text = v.text ?? null;
  } else if (v instanceof Date) {
    out.value = isoOrNull(v);
  } else {
    out.value = v ?? null;
  }
  if (cell.numFmt) out.numFmt = cell.numFmt;
  if (cell.note) out.note = typeof cell.note === 'string' ? cell.note : cell.note.texts?.map(t => t.text).join('') ?? null;
  // Style facets are read back only when the reader materialized them, so a case can
  // assert both survival (a set value comes back) and locality (an unset cell is bare).
  if (cell.fill && cell.fill.type) out.fill = JSON.parse(JSON.stringify(cell.fill));
  if (cell.alignment) out.alignment = JSON.parse(JSON.stringify(cell.alignment));
  return out;
}

export async function roundtripWorkbook(spec) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheets = {};
  for (const s of spec.sheets || []) {
    const sheet = workbook.getWorksheet(s.name);
    if (!sheet) {
      sheets[s.name] = null;
      continue;
    }
    const cells = {};
    for (const c of s.cells || []) cells[c.ref] = normalizeCell(sheet.getCell(c.ref));
    const columns = {};
    for (const col of s.columns || []) {
      const column = sheet.getColumn(col.index);
      columns[col.index] = {
        width: column.width ?? null,
        hidden: !!column.hidden,
        numFmt: column.numFmt ?? null,
      };
    }
    const rows = {};
    for (const row of s.rows || []) {
      const r = sheet.getRow(row.index);
      rows[row.index] = {height: r.height ?? null, hidden: !!r.hidden};
    }
    const ps = sheet.pageSetup || {};
    sheets[s.name] = {
      cells,
      columns,
      rows,
      margins: sheet.pageSetup ? {...sheet.pageSetup.margins} : null,
      pageSetup: {
        fitToPage: !!ps.fitToPage,
        fitToWidth: ps.fitToWidth ?? null,
        fitToHeight: ps.fitToHeight ?? null,
        scale: ps.scale ?? null,
      },
      autoFilter: typeof sheet.autoFilter === 'string' ? sheet.autoFilter : sheet.autoFilter ?? null,
      merges: (sheet.model && sheet.model.merges) || [],
      rowCount: sheet.rowCount,
      actualRowCount: sheet.actualRowCount,
    };
  }
  return {
    properties: {
      creator: workbook.creator ?? null,
      lastModifiedBy: workbook.lastModifiedBy ?? null,
      created: isoOrNull(workbook.created),
      modified: isoOrNull(workbook.modified),
    },
    definedNames: definedNamesOf(workbook),
    sheets,
  };
}

// Read a fixture `.xlsx` and report the workbook-level defined names the reader exposes,
// as { <name>: [ranges…] } — for asserting that a name a real file declares (including a
// full-row/full-column span like `Sheet2!$1:$5`, or two same-named names scoped to different
// sheets) is read back rather than silently dropped by over-strict validation or scope collision.
export async function readFixtureDefinedNames(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const names = definedNamesOf(workbook);
  return {names, count: Object.keys(names).length, modelCount: (workbook.definedNames.model || []).length};
}

// Read the first worksheet of a fixture `.xlsx` with the full (buffered) reader and report the
// requested cells' resolved { type, value } keyed by plain address — for asserting a real file's
// cell values and types (e.g. a Strict-mode ISO-8601 date cell parses to the right date, not a
// spurious 1900-epoch serial). `type` is a stable label; a Date value becomes { date: iso }.
export async function readFixtureCells(rel, cells = []) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const typeName = t => ({2: 'number', 3: 'string', 4: 'date', 6: 'hyperlink', 8: 'richtext'})[t] || `type-${t}`;
  const sheet = workbook.worksheets[0];
  const noteText = n => (n == null ? null : typeof n === 'string' ? n : n.texts?.map(t => t.text).join('') ?? '');
  const out = {};
  for (const addr of cells) {
    const cell = sheet ? sheet.getCell(addr) : null;
    out[addr] = cell
      ? {
          type: typeName(cell.type),
          value: normalizeStreamValue(cell.value),
          numFmt: cell.numFmt ?? null,
          note: cell.note !== undefined ? noteText(cell.note) : undefined,
        }
      : null;
  }
  return out;
}

// Read a fixture `.xlsx` and report the hyperlink cells the reader exposes on the first sheet →
// { <addr>: {hyperlink, text} } — for asserting a real file's hyperlink is reconstructed in full.
// In OOXML the base URL is stored as the relationship target while a `#fragment` is carried
// separately in the hyperlink element's `location`; the reader must rejoin them, not drop the
// fragment (returning the bare base URL). `text` is the display label (rich text flattened).
export async function readFixtureHyperlinks(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const sheet = workbook.worksheets[0];
  const flatten = t => (t == null ? null : typeof t === 'string' ? t : Array.isArray(t.richText) ? t.richText.map(r => r.text).join('') : t);
  const out = {};
  if (sheet) {
    sheet.eachRow({includeEmpty: false}, row => {
      row.eachCell({includeEmpty: false}, cell => {
        const v = cell.value;
        if (v && typeof v === 'object' && 'hyperlink' in v) {
          out[cell.address] = {hyperlink: v.hyperlink ?? null, text: flatten(v.text)};
        }
      });
    });
  }
  return out;
}

// Read a fixture `.xlsx` and report the resolved fill + font of specific cells, keyed
// `<sheet>!<addr>` — for asserting that a real file's cell colors (a solid fill's visible
// foreground color, a theme+tint color, a separate font color) are read back faithfully and
// not conflated. `cells` is a list of `"Sheet!Addr"` strings. Returns plain JSON only.
export async function readFixtureCellStyles(rel, cells = []) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const out = {};
  for (const key of cells) {
    const [sheetName, addr] = key.split('!');
    const sheet = workbook.getWorksheet(sheetName);
    const cell = sheet ? sheet.getCell(addr) : null;
    out[key] = cell
      ? {
          fill: cell.fill && cell.fill.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
          fontColor: cell.font && cell.font.color ? JSON.parse(JSON.stringify(cell.font.color)) : null,
        }
      : null;
  }
  return out;
}

const attrs = tag => {
  const out = {};
  for (const m of String(tag || '').matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = m[2];
  return out;
};

export async function inspectPackage(spec) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const parts = Object.keys(zip.files).filter(f => !zip.files[f].dir).sort();
  const read = async f => (zip.file(f) ? zip.file(f).async('string') : null);

  const contentTypes = (await read('[Content_Types].xml')) || '';
  const workbookXml = (await read('xl/workbook.xml')) || '';
  const relsXml = (await read('xl/_rels/workbook.xml.rels')) || '';

  const worksheetParts = parts.filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
  // First-worksheet relationships: a comment part and a table part must coexist here with
  // unique relationship ids and distinct targets, or Excel repairs the package on open.
  const wsRelsXml = (await read('xl/worksheets/_rels/sheet1.xml.rels')) || '';
  const worksheetRels = [...wsRelsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)].map(t => {
    const a = attrs(t[0]);
    return {id: a.Id, target: a.Target, type: (a.Type || '').split('/').pop()};
  });
  const wsRelIds = worksheetRels.map(r => r.id);
  const overrides = [...contentTypes.matchAll(/<Override[^>]*PartName="([^"]*)"[^>]*\/>/g)].map(m => m[1]);
  // A `<Default>` content-type declaration MUST carry an Extension attribute; a missing one,
  // or a bogus media type like `image/undefined`, produces a package strict Excel repairs.
  const contentTypeDefaults = [...contentTypes.matchAll(/<Default\b[^>]*\/>/g)].map(t => {
    const a = attrs(t[0]);
    return {extension: a.Extension ?? null, contentType: a.ContentType ?? null};
  });
  const sheetEntries = [...workbookXml.matchAll(/<sheet\b[^>]*?\/?>/g)].map(t => {
    const a = attrs(t[0]);
    return {name: a.name, rid: a['r:id']};
  });
  const rels = [...relsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)].map(t => {
    const a = attrs(t[0]);
    return {id: a.Id, target: a.Target, type: (a.Type || '').split('/').pop()};
  });

  const sheets = {};
  const sheetIndex = {};
  (spec.sheets || []).forEach((s, i) => {
    sheetIndex[s.name] = `xl/worksheets/sheet${i + 1}.xml`;
  });
  for (const s of spec.sheets || []) {
    const xml = (await read(sheetIndex[s.name])) || '';
    const marginTag = (xml.match(/<pageMargins\b[^>]*\/>/) || [''])[0];
    const marginAttrs = attrs(marginTag);
    const sheetViewTags = [...xml.matchAll(/<sheetView\b[^>]*(?:\/>|>)/g)];
    const formulas = {};
    for (const m of xml.matchAll(/<c\b[^>]*r="([^"]*)"[^>]*>[\s\S]*?<f\b[^>]*>([\s\S]*?)<\/f>/g)) {
      formulas[m[1]] = m[2];
    }
    // Column groups: a `<col>` addresses columns [min, max]. Excel rejects a file whose
    // col group runs past the sheet's 16384-column limit, so a case asserts on the maxima.
    const columnGroups = [...xml.matchAll(/<col\b[^>]*\/>/g)].map(t => {
      const a = attrs(t[0]);
      return {min: a.min ? Number(a.min) : null, max: a.max ? Number(a.max) : null, width: a.width ?? null};
    });
    // OOXML fixes the order of a worksheet's trailing child elements. In particular the
    // CT_Worksheet sequence requires `drawing` then `legacyDrawing` then `tableParts`; a file
    // that emits `tableParts` before `legacyDrawing` is schema-invalid and Excel repairs it,
    // discarding the sheet. Report the raw positions plus the two adjacency invariants so a case
    // can assert order rather than mere presence (the reader tolerates the wrong order on read).
    const posOf = tag => xml.indexOf(tag);
    const posDrawing = posOf('<drawing ');
    const posLegacy = posOf('<legacyDrawing ');
    const posTable = posOf('<tableParts');
    const ordered = (a, b) => (a >= 0 && b >= 0 ? a < b : null);
    // Header/footer: the `<headerFooter>` element gates its first-page and even-page variants on
    // the `differentFirst`/`differentOddEven` attributes — emitting `<firstHeader>`/`<evenHeader>`
    // without the flag set leaves consuming apps showing the odd content on every page. Report both
    // the child text and the gating flags so a case can assert the flags, not just presence.
    const hfBlock = (xml.match(/<headerFooter\b[\s\S]*?<\/headerFooter>|<headerFooter\b[^>]*\/>/) || [''])[0];
    const hfChild = tag => {
      const m = hfBlock.match(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`));
      return m ? m[1] : null;
    };
    const hfFlag = name => new RegExp(`\\b${name}="(1|true)"`).test(hfBlock);
    sheets[s.name] = {
      pageMargins: {present: Object.keys(marginAttrs), values: marginAttrs},
      hasSheetViews: /<sheetViews>/.test(xml),
      sheetViewCount: sheetViewTags.length,
      hasDimension: /<dimension\b/.test(xml),
      dimensionRef: (xml.match(/<dimension\b[^>]*ref="([^"]*)"/) || [])[1] ?? null,
      autoFilterRef: (xml.match(/<autoFilter\b[^>]*ref="([^"]*)"/) || [])[1] ?? null,
      formulas,
      columnGroups,
      maxColumnIndex: columnGroups.reduce((m, g) => Math.max(m, g.max ?? 0), 0),
      elementOrder: {
        drawing: posDrawing,
        legacyDrawing: posLegacy,
        tableParts: posTable,
        drawingBeforeLegacy: ordered(posDrawing, posLegacy),
        legacyBeforeTableParts: ordered(posLegacy, posTable),
        drawingBeforeTableParts: ordered(posDrawing, posTable),
      },
      headerFooter: {
        present: hfBlock !== '',
        oddHeader: hfChild('oddHeader'),
        oddFooter: hfChild('oddFooter'),
        evenHeader: hfChild('evenHeader'),
        evenFooter: hfChild('evenFooter'),
        firstHeader: hfChild('firstHeader'),
        firstFooter: hfChild('firstFooter'),
        differentOddEven: hfFlag('differentOddEven'),
        differentFirst: hfFlag('differentFirst'),
      },
      xmlWellFormed: xmlWellFormed(xml),
    };
  }

  const tables = [];
  for (const p of parts.filter(f => /^xl\/tables\/table\d+\.xml$/.test(f))) {
    const xml = (await read(p)) || '';
    const a = attrs((xml.match(/<table\b[^>]*>/) || [''])[0]);
    const af = xml.match(/<autoFilter\b[^>]*ref="([^"]*)"/);
    tables.push({
      ref: a.ref ?? null,
      name: a.name ?? null,
      autoFilterRef: af ? af[1] : null,
      columnCount: [...xml.matchAll(/<tableColumn\b/g)].length,
      headerRowCount: a.headerRowCount ?? '1',
      xmlWellFormed: xmlWellFormed(xml),
    });
  }

  // Style facts: a font may reference a color by *theme index* (e.g. <color theme="1"/>),
  // which Excel can only resolve if the package ships a theme part. A theme reference
  // with no theme part is a real corruption mode (Excel repairs the file on open).
  const stylesXml = (await read('xl/styles.xml')) || '';
  const defaultFontBlock = (stylesXml.match(/<font>[\s\S]*?<\/font>/) || [''])[0];
  const defaultFontColor = attrs((defaultFontBlock.match(/<color\b[^>]*\/?>/) || [''])[0]);
  const hasThemePart = parts.some(p => /^xl\/theme\/theme\d+\.xml$/.test(p));
  const styles = {
    hasThemePart,
    defaultFontColor,
    defaultFontUsesTheme: 'theme' in defaultFontColor,
    // The invariant a case locks: any theme-color reference is backed by a theme part.
    themeColorResolvable: !('theme' in defaultFontColor) || hasThemePart,
  };

  // Comment VML: a note's shape is described by a `<v:textbox>` whose inline `style` controls
  // sizing. Without the `mso-fit-shape-to-text:t` directive the host renders every note at a
  // fixed default box that clips multiline text. Surface each textbox style string plus a
  // fit-to-text flag so a case can assert the note box grows to its content instead of clipping.
  const vmlTextboxStyles = [];
  for (const p of parts.filter(f => /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(f))) {
    const vml = (await read(p)) || '';
    for (const t of vml.matchAll(/<(?:v:)?textbox\b[^>]*\bstyle="([^"]*)"/g)) vmlTextboxStyles.push(t[1]);
  }
  const vml = {
    textboxStyles: vmlTextboxStyles,
    allTextboxesFitToText:
      vmlTextboxStyles.length > 0 && vmlTextboxStyles.every(s => /mso-fit-shape-to-text\s*:\s*t/i.test(s)),
  };

  // Worksheet declarations must agree across the three places OOXML requires.
  const declaredConsistent = worksheetParts.every(part => {
    const over = overrides.includes('/' + part);
    const rid = rels.find(r => ('xl/' + r.target).replace('xl/xl/', 'xl/') === part || r.target === part.replace('xl/', ''));
    return over && !!rid;
  });

  return {
    parts,
    worksheetParts,
    overrides,
    contentTypeDefaults,
    sheetEntries,
    rels,
    worksheetRels,
    sheets,
    tables,
    styles,
    vml,
    packageParts: {
      hasCommentsPart: parts.some(p => /^xl\/comments\d+\.xml$/.test(p)),
      hasVmlDrawingPart: parts.some(p => /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(p)),
      hasTablePart: parts.some(p => /^xl\/tables\/table\d+\.xml$/.test(p)),
    },
    consistency: {
      worksheetPartCount: worksheetParts.length,
      sheetEntryCount: sheetEntries.length,
      declaredConsistent,
      worksheetRelIdsUnique: new Set(wsRelIds).size === wsRelIds.length,
    },
  };
}

export async function tryWriteWorkbook(spec) {
  let workbook;
  try {
    workbook = buildFrom(spec);
  } catch (error) {
    return {ok: false, phase: 'build', error: String(error && error.message || error)};
  }
  try {
    const buffer = await workbook.xlsx.writeBuffer();
    // Report which cells survived, so a case can prove a bad cell didn't drop siblings.
    const reread = new ExcelJS.Workbook();
    await reread.xlsx.load(buffer);
    const survivingCells = {};
    for (const s of spec.sheets || []) {
      const sheet = reread.getWorksheet(s.name);
      survivingCells[s.name] = (s.cells || [])
        .filter(c => sheet && sheet.getCell(c.ref).value !== null && sheet.getCell(c.ref).value !== undefined)
        .map(c => c.ref);
    }
    return {ok: true, byteLength: buffer.byteLength ?? buffer.length, survivingCells};
  } catch (error) {
    return {ok: false, phase: 'write', error: String(error && error.message || error)};
  }
}

// Read a fixture `.xlsx` and report the data validations the reader exposes per cell,
// keyed `<sheet>!<address>` → validation model — for asserting that validations a real
// file declares (including ones applied to a multi-cell selection) are read back on
// every cell they cover, not just some.
export async function readFixtureValidations(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const cells = {};
  workbook.eachSheet(sheet => {
    sheet.eachRow({includeEmpty: false}, row => {
      row.eachCell({includeEmpty: false}, cell => {
        if (cell.dataValidation) {
          cells[`${sheet.name}!${cell.address}`] = JSON.parse(JSON.stringify(cell.dataValidation));
        }
      });
    });
  });
  return {cells, count: Object.keys(cells).length};
}

// Read a fixture `.xlsx` and report the DISTINCT data-validation rules each sheet declares,
// read from the worksheet model rather than from populated cells — so a validation applied to
// an otherwise-empty range (a dropdown over a blank column, a COUNTIF guard over future rows)
// is still seen. Rules are de-duplicated by content and reported with how many cells each
// covers, keeping the result bounded even when a rule's sqref nominally spans thousands of
// cells. Lets a case assert that a reference-based list source (a defined name, a cross-sheet
// range) is surfaced verbatim as its formula text rather than stringified to "[object Object]".
export async function readFixtureValidationRules(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const sheets = {};
  workbook.eachSheet(sheet => {
    const model = (sheet.dataValidations && sheet.dataValidations.model) || {};
    const byContent = new Map();
    for (const addr of Object.keys(model)) {
      const rule = model[addr];
      const key = JSON.stringify(rule);
      const entry = byContent.get(key) || {rule: JSON.parse(key), coverageCount: 0};
      entry.coverageCount += 1;
      byContent.set(key, entry);
    }
    sheets[sheet.name] = {rules: [...byContent.values()], ruleCount: byContent.size};
  });
  return {sheets};
}

// Read a fixture `.xlsx`, write it back out, and report the data-validation facts of the
// *re-serialized* package — both standard `<dataValidation>` entries and the extended
// `<x14:dataValidation>` form (2009 extension schema, carried in `<extLst>`, used for
// list validations that reference other sheets or whole columns). Lets a case assert
// that a validation a template declares survives a read→write round-trip rather than
// being silently dropped because the writer only understood the standard form.
export async function roundtripFixtureValidationXml(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const buffer = await workbook.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const parts = Object.keys(zip.files).filter(f => !zip.files[f].dir);
  const sheetParts = parts.filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).sort();

  const sheets = {};
  let totalStandard = 0;
  let totalExt = 0;
  for (const p of sheetParts) {
    const xml = await zip.file(p).async('string');
    // `[ >]` after the tag name distinguishes an individual entry from the
    // `<dataValidations>` / `<x14:dataValidations>` container (next char is `s`).
    const standard = [...xml.matchAll(/<dataValidation[ >]/g)].length;
    const ext = [...xml.matchAll(/<x14:dataValidation[ >]/g)].length;
    const extSqrefs = [...xml.matchAll(/<xm:sqref>([^<]*)<\/xm:sqref>/g)].map(m => m[1]);
    // Per-standard-validation facts, so a case can assert the type, the list/formula source
    // reference, the target range, and the error strings survive the re-serialization intact.
    const standardRules = [...xml.matchAll(/<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g)].map(m => {
      const a = attrs('<x ' + m[1] + '>');
      const f1 = (m[2].match(/<formula1>([\s\S]*?)<\/formula1>/) || [])[1] ?? null;
      return {type: a.type ?? null, sqref: a.sqref ?? null, errorTitle: a.errorTitle ?? null, error: a.error ?? null, formula1: f1};
    });
    sheets[p] = {standardCount: standard, extCount: ext, extSqrefs, standardRules, hasExtLst: /<extLst\b/.test(xml)};
    totalStandard += standard;
    totalExt += ext;
  }
  return {sheets, totalStandard, totalExt, totalValidations: totalStandard + totalExt};
}

// Build a workbook from a spec (cells may carry `formula` or `sharedFormula`), write it,
// read it back, and report each requested cell's resolved formula facts — for asserting
// that a shared-formula clone (a cell that just references a master via `sharedFormula`)
// reads back a concrete, address-translated formula, not an empty cell. Returns per-cell
// { formula, sharedFormula, result } from the model getters (the translated `formula` the
// clone resolves to, distinct from the raw `sharedFormula` master reference).
export async function roundtripFormulas(spec) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const out = {};
  for (const s of spec.sheets || []) {
    const sheet = workbook.getWorksheet(s.name);
    for (const c of s.cells || []) {
      const cell = sheet.getCell(c.ref);
      const v = cell.value;
      out[c.ref] = {
        formula: cell.formula ?? null,
        sharedFormula: v && typeof v === 'object' && 'sharedFormula' in v ? v.sharedFormula : null,
        result: cell.result ?? null,
      };
    }
  }
  return out;
}

// Build a workbook whose sheet carries a table, write it, read it back, then fetch the
// table by name and try to append rows to the *reloaded* table — reporting whether the
// loaded table exposes its rows and whether the append succeeds. Lets a case assert that a
// table rehydrated from a file is mutable, not a half-loaded model that throws on append.
export async function roundtripTableAppend(spec, {tableName, appendRows = []} = {}) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  const table = sheet.getTable(tableName);

  let loadedRowCount = null;
  let loadError = null;
  try {
    const rows = table && table.table ? table.table.rows : table && table.rows;
    loadedRowCount = Array.isArray(rows) ? rows.length : null;
  } catch (e) {
    loadError = String((e && e.message) || e);
  }

  let addError = null;
  let committed = false;
  try {
    for (const r of appendRows) table.addRow(r);
    table.commit();
    committed = true;
  } catch (e) {
    addError = String((e && e.message) || e);
  }

  let finalRowCount = null;
  if (committed) {
    const rewrite = await workbook.xlsx.writeBuffer();
    const reread = new ExcelJS.Workbook();
    await reread.xlsx.load(rewrite);
    const t2 = reread.worksheets[0].getTable(tableName);
    const rows2 = t2 && t2.table ? t2.table.rows : null;
    finalRowCount = Array.isArray(rows2) ? rows2.length : null;
  }

  return {hasTable: !!table, loadedRowCount, loadError, addError, committed, finalRowCount};
}

// Table-part facts extracted from raw table XML — the attributes a roundtrip must not corrupt.
const tableXmlFacts = xml => ({
  hasAutoFilter: /<autoFilter\b/.test(xml),
  autoFilterRef: (xml.match(/<autoFilter\b[^>]*ref="([^"]*)"/) || [])[1] ?? null,
  hasEmptyFilterColumn: /<filterColumn\b[^>]*\/>/.test(xml),
  // headerRowCount defaults to 1 when the attribute is absent; totalsRowShown defaults to 0.
  headerRowCount: (xml.match(/headerRowCount="(\d+)"/) || [])[1] ?? '1',
  totalsRowShown: (xml.match(/totalsRowShown="(\d+)"/) || [])[1] ?? '0',
  name: (xml.match(/<table\b[^>]*\bname="([^"]*)"/) || [])[1] ?? null,
  columnCount: (xml.match(/<tableColumns\b[^>]*count="(\d+)"/) || [])[1] ?? null,
});

const tablesInZip = async zip => {
  const parts = Object.keys(zip.files)
    .filter(f => /^xl\/tables\/table\d+\.xml$/.test(f))
    .sort();
  const facts = [];
  for (const p of parts) facts.push(tableXmlFacts(await zip.file(p).async('string')));
  return facts;
};

// Read a fixture `.xlsx` that contains one or more tables, write it back unchanged, and report
// each table's raw-XML facts before and after — keyed by table name — for asserting that a
// no-op round-trip does not corrupt the table part: it must not inject an autoFilter that was
// not there, flip the header-row configuration off, spuriously turn on totalsRowShown, or emit
// an empty self-closed filterColumn. Excel repairs (and strips) a table whose part is corrupted.
export async function roundtripFixtureTableXml(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const sourceZip = await JSZip.loadAsync(require('node:fs').readFileSync(fixturePath(rel)));
  const source = await tablesInZip(sourceZip);
  const rewrittenZip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const rewritten = await tablesInZip(rewrittenZip);

  // Pair by name where possible; fall back to positional order (a corrupted rewrite can drop
  // the name, so keep the index alignment as a backstop).
  const byName = new Map(rewritten.map(t => [t.name, t]));
  const tables = source.map((s, i) => ({
    name: s.name,
    source: s,
    rewritten: byName.get(s.name) || rewritten[i] || null,
  }));
  return {tables, sourceCount: source.length, rewrittenCount: rewritten.length};
}

// Read a fixture `.xlsx` and report a named table's rehydration facts → { found, columns,
// rowCount } — for asserting that a table loaded from a real file exposes its declared column
// names AND its data rows (populated from the on-sheet cells), not a half-loaded model whose
// rows array is undefined.
export async function readFixtureTable(rel, tableName) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  let table = null;
  workbook.eachSheet(sheet => {
    const t = sheet.getTable(tableName);
    if (t) table = t;
  });
  if (!table) return {found: false, columns: null, rowCount: null};
  const def = table.table || table;
  const columns = Array.isArray(def.columns) ? def.columns.map(c => c.name) : null;
  const rows = def.rows;
  return {found: true, columns, rowCount: Array.isArray(rows) ? rows.length : null};
}

// Read the first worksheet of a fixture `.xlsx` through the STREAMING reader and report the
// requested cells' resolved { type, value } keyed by plain address — for asserting that
// streaming read applies cell styles (so a date-formatted numeric cell is surfaced as a Date,
// matching the full read) rather than leaking the raw serial number. `type` is a stable label.
export async function streamReadFixture(rel, cells = []) {
  const wanted = new Map(cells.map(a => [a, null]));
  const typeName = t => {
    // Map the implementation's ValueType enum to stable labels without leaking the numbers.
    const map = {2: 'number', 3: 'string', 4: 'date', 6: 'hyperlink', 8: 'richtext'};
    return map[t] || `type-${t}`;
  };
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(fixturePath(rel), {});
  for await (const worksheet of reader) {
    for await (const row of worksheet) {
      row.eachCell({includeEmpty: false}, cell => {
        if (wanted.has(cell.address)) wanted.set(cell.address, {type: typeName(cell.type), value: normalizeStreamValue(cell.value)});
      });
    }
    break; // first worksheet only
  }
  const out = {};
  for (const [k, v] of wanted) out[k] = v;
  return out;
}

// Read a fixture `.xlsx` both eagerly and through the STREAMING reader and report the sheet names
// each path surfaces → { eager, streaming } — for asserting the streaming reader joins each
// worksheet part to the workbook-level sheet declaration and exposes the real declared name
// (e.g. "test"), not a generic positional placeholder ("Sheet2"). The eager read is the oracle.
export async function streamVsEagerSheetNames(rel) {
  const eagerWb = new ExcelJS.Workbook();
  await eagerWb.xlsx.readFile(fixturePath(rel));
  const eager = eagerWb.worksheets.map(s => s.name);
  const streaming = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(fixturePath(rel), {});
  for await (const worksheet of reader) streaming.push(worksheet.name);
  return {eager, streaming};
}

// Read a fixture `.xlsx` both eagerly and through the STREAMING reader and report the row numbers
// each path yields for the first sheet → { eager, streaming } — for asserting the streaming reader
// preserves each row's true index across interior blank rows (a sheet with data on rows 1 then
// 6–8 must yield 1,6,7,8, not a resequenced 1,2,3,4) so `row.number` maps back to the real sheet
// position. The eager read is the oracle.
export async function streamVsEagerRowNumbers(rel) {
  const eagerWb = new ExcelJS.Workbook();
  await eagerWb.xlsx.readFile(fixturePath(rel));
  const eager = [];
  eagerWb.worksheets[0].eachRow({includeEmpty: false}, row => eager.push(row.number));
  const streaming = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(fixturePath(rel), {});
  for await (const worksheet of reader) {
    for await (const row of worksheet) streaming.push(row.number);
    break; // first worksheet only
  }
  return {eager, streaming};
}

// Read a fixture `.xlsx` both eagerly and via the STREAMING reader and report each first-sheet
// row's { number, hidden } from both paths → { eager, streaming } — for asserting the streaming
// reader surfaces a row's hidden flag (interpreting the string-form "true"/"false" some generators
// write, not only "1"/"0"), agreeing with the eager read rather than reporting every row visible.
export async function streamVsEagerRowHidden(rel) {
  const eagerWb = new ExcelJS.Workbook();
  await eagerWb.xlsx.readFile(fixturePath(rel));
  const eager = [];
  eagerWb.worksheets[0].eachRow({includeEmpty: false}, row => eager.push({number: row.number, hidden: !!row.hidden}));
  const streaming = [];
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(fixturePath(rel), {});
  for await (const worksheet of reader) {
    for await (const row of worksheet) streaming.push({number: row.number, hidden: !!row.hidden});
    break; // first worksheet only
  }
  return {eager, streaming};
}

// Read a fixture `.xlsx` through the STREAMING reader end-to-end and report { ok, error, sheetNames,
// totalRows } — the read either completes (with every worksheet's name and the total rows delivered
// across sheets) or its error is captured as data. For asserting the streaming reader tolerates a
// package whose ZIP places a worksheet part before xl/workbook.xml (openpyxl and others emit this;
// OOXML does not mandate entry order) rather than crashing on an unbuilt workbook model.
export async function streamReadReport(rel) {
  const sheetNames = [];
  let totalRows = 0;
  try {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(fixturePath(rel), {});
    for await (const worksheet of reader) {
      sheetNames.push(worksheet.name);
      for await (const _row of worksheet) totalRows += 1;
    }
    return {ok: true, error: null, sheetNames, totalRows};
  } catch (e) {
    return {ok: false, error: String((e && e.message) || e), sheetNames, totalRows};
  }
}

const normalizeStreamValue = v => {
  if (v instanceof Date) return {date: Number.isNaN(+v) ? null : v.toISOString()};
  if (v && typeof v === 'object') return JSON.parse(JSON.stringify(v));
  return v ?? null;
};

// Package-part facts a passthrough round-trip must preserve: counts of part families the reader
// does not model (drawings, VML, media, pivot tables/cache, comments) plus worksheet/drawing
// reference flags that wire unmodeled features into the sheet.
async function packageFactsFromZip(zip) {
  const parts = Object.keys(zip.files).filter(f => !zip.files[f].dir);
  const read = async rx => {
    const name = parts.find(p => rx.test(p));
    return name ? await zip.file(name).async('string') : '';
  };
  const ws1 = await read(/worksheets\/sheet1\.xml$/);
  const drawing1 = await read(/drawings\/drawing1\.xml$/);
  return {
    drawings: parts.filter(p => /xl\/drawings\/drawing\d+\.xml$/.test(p)).length,
    vml: parts.filter(p => /vmlDrawing\d+\.vml$/.test(p)).length,
    media: parts.filter(p => /xl\/media\//.test(p)).length,
    pivotTables: parts.filter(p => /pivotTables\/pivotTable\d+\.xml$/.test(p)).length,
    pivotCache: parts.filter(p => /pivotCache\/.+\.xml$/.test(p)).length,
    slicers: parts.filter(p => /slicer/i.test(p)).length,
    comments: parts.filter(p => /comments\d+\.xml$/.test(p)).length,
    hasLegacyDrawingHF: /<legacyDrawingHF\b/.test(ws1),
    hasDrawingRef: /<drawing\b/.test(ws1),
    hasHeaderFooterImageToken: /&amp;G|&G/.test(ws1),
    drawingHasShape: /<xdr:sp\b/.test(drawing1),
    drawingHasPicture: /<xdr:pic\b/.test(drawing1),
  };
}

// Read a fixture `.xlsx`, write it back unchanged, and report package-part facts before/after →
// { source, rewritten } — for asserting a no-op round-trip PRESERVES parts the reader does not
// model (charts, header/footer images and their VML, vector shapes/text boxes, pivot tables and
// their caches) instead of silently dropping them. Each side carries family counts + the
// worksheet/drawing reference flags that wire those features in.
export async function roundtripFixturePackageParts(rel) {
  const buffer = require('node:fs').readFileSync(fixturePath(rel));
  const source = await packageFactsFromZip(await JSZip.loadAsync(buffer));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const rewritten = await packageFactsFromZip(await JSZip.loadAsync(await workbook.xlsx.writeBuffer()));
  return {source, rewritten};
}

// Style-fidelity facts a no-op round-trip must preserve: model-level column widths and pageSetup
// (from the reader), plus raw styles.xml facts (custom indexed-color palette, and the format
// codes of the differential formats used by conditional formatting).
const styleFactsFromZip = async (zip, workbookForModel) => {
  const stylesName = Object.keys(zip.files).find(f => /xl\/styles\.xml$/.test(f));
  const styles = stylesName ? await zip.file(stylesName).async('string') : '';
  const palette = styles.match(/<indexedColors>([\s\S]*?)<\/indexedColors>/);
  const dxfsBlock = (styles.match(/<dxfs\b[\s\S]*?<\/dxfs>/) || [''])[0];
  const sheet = workbookForModel ? workbookForModel.worksheets[0] : null;
  const ps = sheet ? sheet.pageSetup || {} : {};
  return {
    columnWidths:
      sheet && sheet.columns ? sheet.columns.map(c => (c && c.width !== undefined ? c.width : null)).filter(v => v !== null) : [],
    pageSetup: {
      scale: ps.scale ?? null,
      fitToWidth: ps.fitToWidth ?? null,
      fitToHeight: ps.fitToHeight ?? null,
      pageOrder: ps.pageOrder ?? null,
      orientation: ps.orientation ?? null,
    },
    hasIndexedColors: !!palette,
    indexedColorSample: palette ? [...palette[1].matchAll(/rgb="([0-9a-fA-F]+)"/g)].slice(0, 6).map(m => m[1]) : [],
    dxfCount: Number((dxfsBlock.match(/count="(\d+)"/) || [])[1] ?? 0),
    dxfFormatCodes: [...dxfsBlock.matchAll(/<numFmt\b[^>]*formatCode="([^"]*)"/g)].map(m => m[1]),
  };
};

// Read a fixture `.xlsx`, write it back unchanged, and report style-fidelity facts before/after →
// { source, rewritten } — for asserting a no-op round-trip preserves column widths, pageSetup, a
// custom indexed-color palette, and conditional-formatting differential-format number codes
// (which must never serialize as the literal string "[object Object]").
export async function roundtripFixtureStyleFacts(rel) {
  const before = new ExcelJS.Workbook();
  await before.xlsx.readFile(fixturePath(rel));
  const source = await styleFactsFromZip(await JSZip.loadAsync(require('node:fs').readFileSync(fixturePath(rel))), before);
  const buffer = await before.xlsx.writeBuffer();
  // Reload the model for column-width/pageSetup facts, but tolerate a written package the reader
  // chokes on (a corrupt DXF numFmt can make the reload throw) — the raw styles.xml facts still
  // come from the buffer's zip, so a case can assert on what was serialized regardless.
  let after = null;
  try {
    after = new ExcelJS.Workbook();
    await after.xlsx.load(buffer);
  } catch {
    after = null;
  }
  const rewritten = await styleFactsFromZip(await JSZip.loadAsync(buffer), after);
  return {source, rewritten};
}

// Read a fixture `.xlsx`, write it back unchanged, and report the raw serialized `<c>` facts of
// requested cells on the first sheet → { cells: { <addr>: {t, formula, value} }, hasNaNToken } —
// for asserting a load/save round-trip does not corrupt cell content. The guarded failure: a
// string-typed formula cell (t="str") whose style carries a date/number format loses its string
// type on write and its cached result is coerced toward a number, emitting the invalid token
// `NaN` as the cell value — which makes Excel prompt to repair the file on open. `t` is the raw
// cell-type attribute (null when absent → numeric), `formula`/`value` the `<f>`/`<v>` text.
export async function roundtripFixtureCellXml(rel, cells = []) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const name = Object.keys(zip.files).find(f => /xl\/worksheets\/sheet1\.xml$/.test(f));
  const xml = name ? await zip.file(name).async('string') : '';
  const out = {};
  for (const addr of cells) {
    const m = xml.match(new RegExp(`<c\\b[^>]*\\br="${addr}"[^>]*(?:/>|>[\\s\\S]*?</c>)`));
    const el = m ? m[0] : null;
    out[addr] = el
      ? {
          t: (el.match(/\bt="([^"]*)"/) || [])[1] ?? null,
          formula: (el.match(/<f\b[^>]*>([\s\S]*?)<\/f>/) || [])[1] ?? null,
          value: (el.match(/<v\b[^>]*>([\s\S]*?)<\/v>/) || [])[1] ?? null,
        }
      : null;
  }
  // A numeric cell whose text is the literal `NaN` is invalid OOXML content — scan the whole sheet.
  const hasNaNToken = /<v[^>]*>\s*NaN\s*<\/v>/.test(xml);
  return {cells: out, hasNaNToken};
}

// Conditional-formatting rule facts from a worksheet's raw XML: every cfRule with its type,
// dxfId, and priority, plus the count of conditionalFormatting blocks. A rule type the library
// does not otherwise model (e.g. duplicateValues) must still survive a round-trip for I/O
// fidelity — dropping it (or emptying its conditionalFormatting shell) makes Excel repair the file.
const cfFactsFromXml = xml => ({
  blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
  rules: [...xml.matchAll(/<cfRule\b([^>]*?)\/?>/g)].map(m => {
    const a = attrs('<x ' + m[1] + '>');
    return {type: a.type ?? null, dxfId: a.dxfId ?? null, priority: a.priority ?? null};
  }),
});

// Read a fixture `.xlsx`, write it back unchanged, and report conditional-formatting facts of the
// first sheet before/after → { source, rewritten } — for asserting a no-op round-trip preserves a
// conditional-formatting rule (its type, dxfId, and priority) rather than dropping it or emitting
// an empty conditionalFormatting shell with no cfRule, which corrupts the file.
export async function roundtripFixtureConditionalFormatting(rel) {
  const srcZip = await JSZip.loadAsync(require('node:fs').readFileSync(fixturePath(rel)));
  const srcName = Object.keys(srcZip.files).find(f => /xl\/worksheets\/sheet1\.xml$/.test(f));
  const source = cfFactsFromXml(srcName ? await srcZip.file(srcName).async('string') : '');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const outZip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const outName = Object.keys(outZip.files).find(f => /xl\/worksheets\/sheet1\.xml$/.test(f));
  const rewritten = cfFactsFromXml(outName ? await outZip.file(outName).async('string') : '');
  return {source, rewritten};
}

// Read a fixture `.xlsx`, write it back unchanged, reload it, and report how many styled cells'
// VISIBLE fill and border colors changed across the round-trip → { checked, fillMismatches,
// borderMismatches, fillSample, borderSample }. A benign `patternFill pattern="none"` that the
// writer adds to an unfilled cell is ignored (it renders identically) — only a real fill color
// (a solid/patterned fgColor) or a border-edge color that diverges counts, so a case can lock
// that themed/indexed fill and border colors survive a pure open-then-save.
export async function roundtripFixtureColorFidelity(rel) {
  const before = new ExcelJS.Workbook();
  await before.xlsx.readFile(fixturePath(rel));
  const after = new ExcelJS.Workbook();
  await after.xlsx.load(await before.xlsx.writeBuffer());

  const realFill = cell => (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern !== 'none' ? cell.fill : null);
  const borderColors = cell => {
    if (!cell.border) return null;
    const out = {};
    for (const edge of ['top', 'left', 'right', 'bottom']) {
      if (cell.border[edge] && cell.border[edge].color) out[edge] = cell.border[edge].color;
    }
    return Object.keys(out).length ? out : null;
  };
  const norm = v => JSON.stringify(stableSort(v ?? null));

  let checked = 0;
  let fillMismatches = 0;
  let borderMismatches = 0;
  let fillSample = null;
  let borderSample = null;
  before.eachSheet(sheet => {
    const other = after.getWorksheet(sheet.name);
    sheet.eachRow({includeEmpty: false}, row => {
      row.eachCell({includeEmpty: false}, cell => {
        if (!realFill(cell) && !borderColors(cell)) return;
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
      });
    });
  });
  return {checked, fillMismatches, borderMismatches, fillSample, borderSample};
}

// The value of the print-area defined name (`_xlnm.Print_Area`) in a workbook.xml, split into its
// comma-separated ranges. Excel records multiple print areas on one sheet as ONE Print_Area name
// whose value is a comma-separated range list sharing one localSheetId.
const printAreaRanges = workbookXml => {
  const m = workbookXml.match(/<definedName\b[^>]*name="_xlnm\.Print_Area"[^>]*>([\s\S]*?)<\/definedName>/);
  if (!m) return [];
  // The value carries XML entity escapes (&apos; around a quoted sheet name); ranges are comma-split
  // at the top level (range refs contain no commas).
  return m[1].replace(/&apos;/g, "'").split(',').map(s => s.trim()).filter(Boolean);
};

// Read a fixture `.xlsx` that declares more than one print area on a sheet, and report the print
// areas along three axes → { sourceRangeCount, readPrintArea, rewrittenRangeCount } — for asserting
// that a single Print_Area defined name holding a comma-separated list of ranges is fully recovered
// (both ranges, not just the first) and re-emitted as both ranges, rather than being truncated to
// one on read and mangled on write.
export async function roundtripFixturePrintAreas(rel) {
  const srcZip = await JSZip.loadAsync(require('node:fs').readFileSync(fixturePath(rel)));
  const srcWbName = Object.keys(srcZip.files).find(f => /xl\/workbook\.xml$/.test(f));
  const sourceRangeCount = printAreaRanges(srcWbName ? await srcZip.file(srcWbName).async('string') : '').length;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const readPrintArea = workbook.worksheets[0].pageSetup ? workbook.worksheets[0].pageSetup.printArea ?? null : null;

  const outZip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const outWbName = Object.keys(outZip.files).find(f => /xl\/workbook\.xml$/.test(f));
  const rewrittenRangeCount = printAreaRanges(outWbName ? await outZip.file(outWbName).async('string') : '').length;

  return {sourceRangeCount, readPrintArea, rewrittenRangeCount};
}

// Build a workbook whose one sheet declares a print area (a `printArea` string, which may be a
// comma-separated list of ranges for multiple print areas), write it, and report the ranges of the
// emitted `_xlnm.Print_Area` defined name → { rangeCount, ranges } — for asserting that authoring
// two print areas produces one sheet-scoped defined name carrying both ranges, not a single
// truncated/mangled one.
export async function writePrintAreaDefinedName(printArea) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('S');
  sheet.getCell('A1').value = 1;
  sheet.pageSetup.printArea = printArea;
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
  const wbName = Object.keys(zip.files).find(f => /xl\/workbook\.xml$/.test(f));
  const ranges = printAreaRanges(wbName ? await zip.file(wbName).async('string') : '');
  return {rangeCount: ranges.length, ranges};
}

// Read a fixture `.xlsx` and report only { ok, error, sheetNames } — for asserting the
// reader neither crashes nor mis-reads a workbook produced by a *foreign* (non-Excel)
// generator: a namespace-prefixed OOXML root (`<x:workbook>` instead of `<workbook>`),
// a leading byte-order mark before the XML declaration, a non-ASCII sheet name, or parts
// ordered unusually within the zip. The read error is captured and returned as data
// (never propagated) so a case can assert on a crash rather than the runner blowing up.
export async function readFixtureReport(rel) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(fixturePath(rel));
    return {ok: true, error: null, sheetNames: workbook.worksheets.map(s => s.name)};
  } catch (e) {
    return {ok: false, error: String((e && e.message) || e), sheetNames: null};
  }
}

// Recursively sort object keys so two style objects with the same content but different
// key order compare equal — the reader emits font/fill fields in load-dependent order.
const stableSort = v => {
  if (Array.isArray(v)) return v.map(stableSort);
  if (v && typeof v === 'object') {
    const out = {};
    for (const k of Object.keys(v).sort()) out[k] = stableSort(v[k]);
    return out;
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
  wb.eachSheet(sheet => {
    const cols = {};
    sheet.columns.forEach((col, i) => {
      if (col && col.width !== undefined) cols[i + 1] = {width: col.width, customWidth: !!col.isCustomWidth};
    });
    out[sheet.name] = cols;
  });
  return out;
};

// Read a fixture, write it back unchanged, and read the result again — then report whether
// the styling a real template declares survives that no-op read→write→read round-trip:
// sheet names, custom column widths, and per-cell fill/font/numFmt/border/alignment. This
// is the mainstream "open a styled template, fill it in, save it" path, which must be
// format-preserving. Style comparison is key-order-insensitive so a case asserts on
// content survival, not serialization incidentals.
export async function roundtripFixture(rel) {
  const before = new ExcelJS.Workbook();
  await before.xlsx.readFile(fixturePath(rel));
  const buffer = await before.xlsx.writeBuffer();
  const after = new ExcelJS.Workbook();
  await after.xlsx.load(buffer);

  let checked = 0;
  let mismatches = 0;
  let sample = null;
  before.eachSheet(sheet => {
    const other = after.getWorksheet(sheet.name);
    sheet.eachRow({includeEmpty: false}, row => {
      row.eachCell({includeEmpty: false}, cell => {
        if (!hasStyle(cell)) return;
        checked += 1;
        const beforeKey = styleKey(cell);
        const afterKey = other ? styleKey(other.getCell(cell.address)) : '(sheet missing)';
        if (beforeKey !== afterKey) {
          mismatches += 1;
          if (!sample) sample = {cell: `${sheet.name}!${cell.address}`, before: beforeKey, after: afterKey};
        }
      });
    });
  });

  return {
    sheetNamesBefore: before.worksheets.map(s => s.name),
    sheetNames: after.worksheets.map(s => s.name),
    columnsBefore: columnsWithWidth(before),
    columns: columnsWithWidth(after),
    styleSurvival: {checked, mismatches, sample},
  };
}

const intAt = (xml, tag) => {
  const m = xml.match(new RegExp(`<${tag}>(-?\\d+)</${tag}>`));
  return m ? Number(m[1]) : null;
};

const parseAnchorSide = block =>
  block
    ? {
        col: intAt(block, 'xdr:col'),
        colOff: intAt(block, 'xdr:colOff'),
        row: intAt(block, 'xdr:row'),
        rowOff: intAt(block, 'xdr:rowOff'),
      }
    : null;

// Build a workbook whose sheets place images at given ranges (see the `images` field of
// the sheet spec), write it, unzip, and report the drawing-anchor geometry that was
// actually serialized — for asserting that a fractional/whole anchor coordinate maps to
// the correct OOXML col/colOff/row/rowOff against the *real* column width and row height,
// and that a string range becomes a valid two-cell anchor. Everything returned is plain
// numbers, never an implementation object (the live anchor holds a back-reference to the
// worksheet).
export async function inspectImageAnchors(spec) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const drawingParts = Object.keys(zip.files)
    .filter(f => /^xl\/drawings\/drawing\d+\.xml$/.test(f))
    .sort();

  const anchors = [];
  let xmlOk = true;
  for (const p of drawingParts) {
    const xml = await zip.file(p).async('string');
    if (!xmlWellFormed(xml)) xmlOk = false;
    for (const m of xml.matchAll(/<xdr:(oneCellAnchor|twoCellAnchor)\b([^>]*)>([\s\S]*?)<\/xdr:\1>/g)) {
      const body = m[3];
      const fromBlock = (body.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/) || [])[1];
      const toBlock = (body.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/) || [])[1];
      const extTag = body.match(/<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
      const editAs = (m[2].match(/editAs="([^"]*)"/) || [])[1] || null;
      // The picture's own shape transform (spPr/xfrm). Excel ignores it for an anchored drawing and
      // positions purely from the anchor, but a strict consumer (LibreOffice) honors it — so a
      // zeroed placeholder transform (off 0,0 + ext 0,0) detaches the image from its anchor cell.
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
          // The interop hazard: a transform present but zeroed, which overrides the anchor.
          zeroedTransform: !!(off && spExt && off.x === 0 && off.y === 0 && spExt.cx === 0 && spExt.cy === 0),
        },
      });
    }
  }
  return {anchors, drawingCount: drawingParts.length, xmlWellFormed: xmlOk};
}

// Read a fixture `.xlsx` and report each image's normalized anchor range — for asserting
// that a file whose drawing anchors were authored as cell ranges (including string ranges
// like "B198:BN198") reads without crashing and exposes an object range with integer
// cell coordinates, never a raw string or a throw. Returns plain numbers only.
export async function readFixtureImageAnchors(rel) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(fixturePath(rel));
  const images = [];
  workbook.eachSheet(sheet => {
    for (const im of sheet.getImages()) {
      const r = im.range || {};
      const side = a => (a ? {col: a.nativeCol ?? null, row: a.nativeRow ?? null} : null);
      images.push({sheet: sheet.name, editAs: r.editAs ?? null, tl: side(r.tl), br: side(r.br)});
    }
  });
  return {images, count: images.length};
}

// Normalize a CSV-read cell value to plain JSON: a Date becomes { date: iso }, an Excel
// error object stays { error }, empties are null, primitives pass through — so a case can
// distinguish "read as a string" from "coerced to a Date/number" without holding a Date.
const normalizeCsvValue = v => {
  if (v instanceof Date) return {date: Number.isNaN(+v) ? null : v.toISOString()};
  if (v && typeof v === 'object' && 'error' in v) return {error: v.error};
  return v ?? null;
};

// Parse a CSV *string* through the reader with the given options and report
// { ok, error, rows } where rows is a JSON-serializable 2-D array of typed cell values —
// for asserting delimiter handling, value coercion (numbers/dates/identifiers), and that
// a broken option path surfaces as a captured error rather than crashing the runner. The
// legacy `map`/`parserOptions.map` is a function and cannot be expressed as JSON, so cases
// drive behavior via `parserOptions`/`dateFormats`, not custom callbacks.
export async function csvRead({csv, options} = {}) {
  const workbook = new ExcelJS.Workbook();
  try {
    const worksheet = await workbook.csv.read(Readable.from([csv]), options);
    const rows = [];
    worksheet.eachRow({includeEmpty: true}, row => {
      rows.push(row.values.slice(1).map(normalizeCsvValue));
    });
    return {ok: true, error: null, rows};
  } catch (e) {
    return {ok: false, error: String((e && e.message) || e), rows: []};
  }
}

const csvCellValue = c => {
  if (c && typeof c === 'object') {
    if (c.date) return toDate(c.date);
    if ('formula' in c) return {formula: c.formula, result: c.result};
    if ('error' in c) return {error: c.error};
  }
  return c;
};

// Build a worksheet from a declarative `{ rows: [[cell,…]] }` spec (a cell is a primitive,
// { date: iso }, { formula, result }, or { error }), write it to CSV with the given
// options, and report { ok, error, text }. Lets a case assert on the produced CSV text —
// field delimiter, date formatting — for genuinely-typed cells (a real Date, not a string).
export async function csvWrite({spec = {}, options} = {}) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('S');
  for (const row of spec.rows || []) worksheet.addRow((row || []).map(csvCellValue));
  try {
    const buffer = await workbook.csv.writeBuffer(options);
    return {ok: true, error: null, text: buffer.toString().replace(/\r?\n$/, '')};
  } catch (e) {
    return {ok: false, error: String((e && e.message) || e), text: null};
  }
}

// Materialize a stream-write cell spec into a live cell value: a { richText:[{text,
// bold?, italic?}] } spec becomes a real richText value; primitives pass through.
const streamCellValue = c => {
  if (c && typeof c === 'object' && Array.isArray(c.richText)) {
    return {
      richText: c.richText.map(run => ({
        text: run.text,
        font: {...(run.bold ? {bold: true} : {}), ...(run.italic ? {italic: true} : {})},
      })),
    };
  }
  return c;
};

// Read a richText (or primitive) cell value back into a plain shape a case can compare:
// { richText:[{text, bold, italic}] } or the primitive itself.
const readStreamCell = v => {
  if (v && typeof v === 'object' && Array.isArray(v.richText)) {
    return {
      richText: v.richText.map(run => ({
        text: run.text,
        bold: !!(run.font && run.font.bold),
        italic: !!(run.font && run.font.italic),
      })),
    };
  }
  return v ?? null;
};

// Drive the streaming workbook writer through a sequence of row ops, commit, then read the
// produced package back and report the resulting cells — for asserting streaming-write
// behavior (batch add, shared-string handling of richText) that the non-streaming path
// can't exercise. Ops: { op:'addRow', value:[…] } | { op:'addRows', value:[[…],…] }. A row
// op that throws (e.g. an unimplemented method) is reported as { ok:false, error }, never
// propagated. Returns { ok, error, cells:{ <ref>: value }, rowCount } with plain values.
export async function streamWriteSheet({useSharedStrings = false, ops = [], read = []} = {}) {
  const chunks = [];
  const stream = new PassThrough();
  const drained = new Promise(res => {
    stream.on('data', c => chunks.push(c));
    stream.on('end', res);
    stream.on('close', res);
  });
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream, useSharedStrings});
  let error = null;
  try {
    const sheet = writer.addWorksheet('S');
    for (const op of ops) {
      if (op.op === 'addRow') {
        sheet.addRow((op.value || []).map(streamCellValue)).commit();
      } else if (op.op === 'addRows') {
        sheet.addRows((op.value || []).map(row => (row || []).map(streamCellValue)));
      } else {
        throw new Error(`unknown stream op: ${op.op}`);
      }
    }
    sheet.commit();
  } catch (e) {
    error = String((e && e.message) || e);
  }
  await writer.commit();
  await drained;

  if (error) return {ok: false, error, cells: {}, rowCount: 0};

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.concat(chunks));
  const sheet = workbook.worksheets[0];
  const cells = {};
  for (const ref of read) cells[ref] = readStreamCell(sheet.getCell(ref).value);
  return {ok: true, error: null, cells, rowCount: sheet.rowCount};
}

// Drive the streaming writer to produce a whole package, then treat the emitted bytes as an
// untrusted archive and report its container integrity: every declared zip entry is present with
// a non-zero byte length, every entry's stored CRC matches its decompressed bytes (JSZip verifies
// CRC-32 on extraction when asked), and the package re-reads to the same sheet names and cell
// values a whole-file write would yield. Lets a case assert that the streaming path assembles a
// valid zip — not merely valid XML — because a bad CRC or a zero-byte auxiliary part makes Excel
// reject the file even when the sheet XML is perfect.
export async function streamWritePackageReport({rows = 50} = {}) {
  const chunks = [];
  const stream = new PassThrough();
  const drained = new Promise(res => {
    stream.on('data', c => chunks.push(c));
    stream.on('end', res);
    stream.on('close', res);
  });
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
  const sheet = writer.addWorksheet('S');
  for (let i = 1; i <= rows; i++) sheet.addRow([`r${i}`, i]).commit();
  sheet.commit();
  await writer.commit();
  await drained;

  const buffer = Buffer.concat(chunks);
  let crcValid = true;
  let crcError = null;
  const emptyParts = [];
  let partCount = 0;
  try {
    const zip = await JSZip.loadAsync(buffer, {checkCRC32: true});
    const names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
    partCount = names.length;
    for (const n of names) {
      // Extracting with checkCRC32 on throws if the stored CRC disagrees with the bytes.
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
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    sheetNames = wb.worksheets.map(s => s.name);
    const s = wb.worksheets[0];
    for (let i = 1; i <= Math.min(rows, 3); i++) firstCol.push(readStreamCell(s.getCell(`A${i}`).value));
  } catch (e) {
    reloadOk = false;
    reloadError = String((e && e.message) || e);
  }

  return {partCount, emptyParts, crcValid, crcError, reloadOk, reloadError, sheetNames, firstCol};
}

// Write a declarative `spec` to a package, then read it back through the STREAMING reader (fed the
// bytes as a Readable so the SAX parser sees real chunk boundaries) and return the requested cell
// values → { <ref>: value }. Pairs with a whole-file read of the same spec to assert the streaming
// path is byte-exact — in particular that multi-byte UTF-8 (CJK / emoji) split across a chunk
// boundary is reassembled rather than decoded per-chunk into U+FFFD replacement characters.
export async function streamReadSpec(spec, cells = []) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const wanted = new Set(cells);
  const streamed = {};
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from(Buffer.from(buffer)), {});
  for await (const worksheet of reader) {
    for await (const row of worksheet) {
      row.eachCell({includeEmpty: false}, cell => {
        if (wanted.has(cell.address)) streamed[cell.address] = normalizeStreamValue(cell.value);
      });
    }
    break; // first worksheet only
  }
  const eager = new ExcelJS.Workbook();
  await eager.xlsx.load(buffer);
  const es = eager.worksheets[0];
  const eagerCells = {};
  for (const ref of cells) eagerCells[ref] = normalizeStreamValue(es.getCell(ref).value);
  return {streamed, eager: eagerCells};
}

// Author a workbook where several cells share one on-disk style index (identical formatting is
// deduplicated to a single style record), write it, load it back, then mutate ONE loaded cell's
// style property and read a SIBLING cell's same property. Reports whether the mutation bled across
// cells → { sibling, mutatedTo, bled } — for asserting the reader hands each loaded cell an
// INDEPENDENT style object rather than aliasing the shared record so one edit corrupts the rest.
export async function loadMutateCellStyle({sharedFill = 'FFFF0000', mutateTo = 'FF00FF00'} = {}) {
  const wb = new ExcelJS.Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'a';
  s.getCell('B1').value = 'b';
  const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: sharedFill}};
  s.getCell('A1').fill = fill;
  s.getCell('B1').fill = fill; // identical formatting → one shared style index on disk
  const buffer = await wb.xlsx.writeBuffer();

  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer);
  const s2 = wb2.getWorksheet('S');
  s2.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: mutateTo}};
  const siblingFg =
    s2.getCell('B1').fill && s2.getCell('B1').fill.fgColor ? s2.getCell('B1').fill.fgColor.argb : null;

  // Write the mutated workbook back and confirm only the one cell changed on disk.
  const rewrite = await wb2.xlsx.writeBuffer();
  const wb3 = new ExcelJS.Workbook();
  await wb3.xlsx.load(rewrite);
  const s3 = wb3.getWorksheet('S');
  const diskSiblingFg =
    s3.getCell('B1').fill && s3.getCell('B1').fill.fgColor ? s3.getCell('B1').fill.fgColor.argb : null;

  return {
    sibling: siblingFg,
    mutatedTo: mutateTo,
    original: sharedFill,
    bled: siblingFg === mutateTo,
    diskSibling: diskSiblingFg,
    diskBled: diskSiblingFg === mutateTo,
  };
}

// Copy one worksheet onto another via the worksheet `model` export/import contract (the idiomatic
// "clone a sheet" pattern) and report whether merged ranges survive → { srcMerges, dstMerges,
// error }. Isolates the model serialize/deserialize symmetry: a source sheet's merges must reappear
// on the destination after `dst.model = {...src.model, name}`, not silently vanish.
export function copyWorksheetModel({merges = ['A1:C1'], cells = [{ref: 'A1', value: 'merged'}]} = {}) {
  const wb = new ExcelJS.Workbook();
  const src = wb.addWorksheet('Src');
  for (const c of cells) src.getCell(c.ref).value = c.value;
  for (const m of merges) src.mergeCells(m);
  const dst = wb.addWorksheet('Dst');
  let error = null;
  let dstMerges = [];
  const srcMerges = (src.model && src.model.merges) || [];
  try {
    dst.model = {...src.model, name: 'Dst'};
    dstMerges = (dst.model && dst.model.merges) || [];
  } catch (e) {
    error = String((e && e.message) || e);
  }
  return {srcMerges: [...srcMerges].sort(), dstMerges: [...dstMerges].sort(), error};
}

// Build a workbook from a spec, write it, and report the written style table's size plus the
// style index each requested cell resolved to → { cellXfCount, indices: { <ref>: index|null } }.
// styles.xml is meant to be a SHARED table referenced by index: many cells carrying identical
// visual formatting must collapse to one cellXfs entry (and one shared index), while a genuinely
// different style stays a distinct entry — so a case can assert dedup neither fails (one entry per
// cell, the historical write-time cliff) nor over-collapses distinct styles. A cell left at the
// default style carries no `s` attribute and reports null.
export async function styleDedupReport(spec, cells = []) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const styles = (await zip.file('xl/styles.xml').async('string')) || '';
  const xfBlock = (styles.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [''])[0];
  const cellXfCount = (xfBlock.match(/<xf\b/g) || []).length;
  const sheetXml = (await zip.file('xl/worksheets/sheet1.xml').async('string')) || '';
  const indices = {};
  for (const ref of cells) {
    const m = sheetXml.match(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*\\bs="(\\d+)"`));
    indices[ref] = m ? Number(m[1]) : null;
  }
  return {cellXfCount, indices};
}

// Build a workbook from a spec, write it, load it back, and report — per requested row — which
// column indices a full (`includeEmpty`) cell iteration yields, plus the row's own cell/values
// length and the sheet's declared column count → { rows: { <n>: { cols, cellCount, valuesLength } },
// columnCount }. For asserting positional row reconstruction: a row that stops short of the sheet's
// last populated column should still surface its trailing empty cells so every data row aligns
// column-for-column with a wider header, and interior vs. trailing empties are treated consistently.
export async function readRowCellPresence(spec, rows = []) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.worksheets[0];
  const out = {};
  for (const rn of rows) {
    const row = sheet.getRow(rn);
    const cols = [];
    row.eachCell({includeEmpty: true}, (_cell, col) => cols.push(col));
    out[rn] = {cols, cellCount: row.cellCount, valuesLength: row.values.length};
  }
  return {rows: out, columnCount: sheet.columnCount};
}

// Build a workbook from a spec, write it, and read the requested rows' `values` arrays both eagerly
// and through the STREAMING reader → { eager: { <n>: [...] }, streamed: { <n>: [...] } }. Sparse
// holes normalize to null so the JSON is comparable. Columns are 1-based, so a row's `values[0]` is
// an empty leading slot and the first real cell lands at index 1; the durable invariant is that the
// streaming reader exposes the SAME indexing convention as the full-load reader, so a caller can
// switch read modes without re-indexing.
export async function streamVsEagerRowValues(spec, rowNumbers = [1]) {
  const buffer = await buildFrom(spec).xlsx.writeBuffer();
  const holes = arr => Array.from(arr, v => (v === undefined ? null : normalizeStreamValue(v)));
  const eagerWb = new ExcelJS.Workbook();
  await eagerWb.xlsx.load(buffer);
  const es = eagerWb.worksheets[0];
  const eager = {};
  for (const n of rowNumbers) eager[n] = holes(es.getRow(n).values);
  const streamed = {};
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from(Buffer.from(buffer)), {});
  for await (const worksheet of reader) {
    for await (const row of worksheet) {
      if (rowNumbers.includes(row.number)) streamed[row.number] = holes(row.values);
    }
    break; // first worksheet only
  }
  return {eager, streamed};
}

// Build a workbook whose sheet carries one or more tables from a spec, write it (facts A), then load
// it and write it again (facts B) → { write, roundtrip, loadOk, loadError }, each fact list carrying
// per-table { ref, name, wellFormed }. For asserting a defined table survives a load→save round-trip
// with its reference range intact and its table part not dropped — the fidelity mainstream files
// depend on — including the degenerate empty-body and single-data-row shapes that must not error.
export async function roundtripSpecTableFacts(spec) {
  const tableFacts = async zip => {
    const parts = Object.keys(zip.files).filter(f => /^xl\/tables\/table\d+\.xml$/.test(f)).sort();
    const facts = [];
    for (const p of parts) {
      const xml = await zip.file(p).async('string');
      facts.push({
        ref: (xml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
        name: (xml.match(/<table\b[^>]*\bname="([^"]*)"/) || [])[1] ?? null,
        wellFormed: xmlWellFormed(xml),
      });
    }
    return facts;
  };
  const writeBuf = await buildFrom(spec).xlsx.writeBuffer();
  const write = await tableFacts(await JSZip.loadAsync(writeBuf));
  let loadOk = true;
  let loadError = null;
  let roundtrip = [];
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(writeBuf);
    const rewrite = await wb.xlsx.writeBuffer();
    roundtrip = await tableFacts(await JSZip.loadAsync(rewrite));
  } catch (e) {
    loadOk = false;
    loadError = String((e && e.message) || e);
  }
  return {write, roundtrip, loadOk, loadError};
}

// Author two cells sharing one font, write, load, then reassign ONE loaded cell's font by SPREADING
// its existing font and overriding a member (`cell.font = {...cell.font, color}` — the idiomatic
// "tweak one property" pattern), and read the sibling's font color → { edited, sibling, bled }. The
// companion to loadMutateCellStyle for the font facet: even the spread-then-override path (which
// builds a fresh object literal) must not mutate the sibling that shared the on-disk style record.
export async function loadMutateCellFont({original = 'FF000000', mutateTo = 'FFFF0000'} = {}) {
  const wb = new ExcelJS.Workbook();
  const s = wb.addWorksheet('S');
  const font = {name: 'Arial', size: 12, color: {argb: original}};
  s.getCell('A1').value = 'a';
  s.getCell('A1').font = font;
  s.getCell('B1').value = 'b';
  s.getCell('B1').font = font; // identical formatting → one shared style index on disk
  const buffer = await wb.xlsx.writeBuffer();

  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buffer);
  const s2 = wb2.getWorksheet('S');
  s2.getCell('A1').font = {...s2.getCell('A1').font, color: {argb: mutateTo}};
  const colorOf = cell => (cell.font && cell.font.color ? cell.font.color.argb : null);
  const sibling = colorOf(s2.getCell('B1'));
  return {edited: colorOf(s2.getCell('A1')), sibling, original, mutatedTo: mutateTo, bled: sibling === mutateTo};
}

function xmlWellFormed(xml) {
  // Cheap structural check: no raw & that isn't an entity, tags balanced enough to
  // matter. A real parser lives in the impl; here we only need "would a strict
  // consumer choke", which unescaped & or < inside text triggers.
  if (/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml)) return false;
  return true;
}

// Author list-type data validations programmatically — an inline quoted literal
// ("Male,Female") and a cross-sheet range reference (Levels!$A$2:$A$9999) — write the
// workbook, and report both what the reader hands back per cell and the serialized
// `<dataValidations>` facts. Lets a case assert that BOTH the value-source forms a real
// author uses survive a write→read round-trip and that the emitted XML is well-formed and
// count-correct, without the case knowing how validations are shaped internally.
export async function authorListValidations(validations = []) {
  const workbook = new ExcelJS.Workbook();
  const main = workbook.addWorksheet('Main');
  const levels = workbook.addWorksheet('Levels');
  levels.getCell('A2').value = 'X';
  for (const v of validations) {
    const dv = {type: 'list', allowBlank: v.allowBlank !== false, formulae: [v.formula]};
    if (v.error !== undefined) {
      dv.showErrorMessage = true;
      dv.error = v.error;
    }
    main.getCell(v.ref).dataValidation = dv;
  }
  const buf = await workbook.xlsx.writeBuffer();

  const reread = new ExcelJS.Workbook();
  await reread.xlsx.load(buf);
  const sheet = reread.getWorksheet('Main');
  const readBack = {};
  for (const v of validations) {
    const dv = sheet.getCell(v.ref).dataValidation;
    readBack[v.ref] = dv ? {type: dv.type, formulae: dv.formulae ?? null} : null;
  }

  const zip = await JSZip.loadAsync(buf);
  const name = Object.keys(zip.files).find(n => /sheet1\.xml$/.test(n));
  const xml = name ? await zip.files[name].async('string') : '';
  const block = (xml.match(/<dataValidations[\s\S]*?<\/dataValidations>/) || [])[0] || '';
  return {
    readBack,
    xml: {
      count: [...xml.matchAll(/<dataValidation[ >]/g)].length,
      wellFormed: xmlWellFormed(block),
      // The exact serialized formula1 texts, so a case can assert the quoting/reference
      // survived verbatim rather than being mangled or coerced.
      formula1: [...block.matchAll(/<formula1>([\s\S]*?)<\/formula1>/g)].map(m =>
        m[1].replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')
      ),
    },
  };
}

// Author per-cell protection flags plus a protected sheet, write, and report (a) what the
// reader hands back for each cell's protection, (b) whether the cell style record actually
// carries the protection (applyProtection + <protection> in cellXfs), and (c) the
// `<sheetProtection>` element the protected sheet emits. Lets a case assert the meaningful
// case — an *unlocked* cell survives (default is locked, so only unlocked is distinguishable)
// — and that worksheet protection, the thing that makes locked flags enforceable, is emitted.
export async function authorCellProtection(cells = [], protect = null) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('S');
  for (const c of cells) {
    sheet.getCell(c.ref).value = c.value ?? c.ref;
    if (c.protection !== undefined) sheet.getCell(c.ref).protection = c.protection;
  }
  if (protect) await sheet.protect(protect.password ?? undefined, protect.options ?? {});
  const buf = await workbook.xlsx.writeBuffer();

  const reread = new ExcelJS.Workbook();
  await reread.xlsx.load(buf);
  const sheet2 = reread.getWorksheet('S');
  const readBack = {};
  for (const c of cells) {
    const p = sheet2.getCell(c.ref).protection;
    readBack[c.ref] = p ? {locked: p.locked ?? null} : null;
  }

  const zip = await JSZip.loadAsync(buf);
  const styles = await zip.files['xl/styles.xml'].async('string');
  const sheetName = Object.keys(zip.files).find(n => /sheet1\.xml$/.test(n));
  const sheetXml = sheetName ? await zip.files[sheetName].async('string') : '';
  return {
    readBack,
    hasApplyProtection: /applyProtection="1"/.test(styles) && /<protection\b/.test(styles),
    sheetProtection: (sheetXml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0] || null,
  };
}

// Drive the streaming workbook writer over a CALLER-SUPPLIED writable stream (a plain
// PassThrough or a Duplex) rather than a library-owned file stream, and report whether the
// workbook-commit promise settles within a bounded time plus the byte total collected from
// the sink. Streaming straight to a remote sink (object storage, an HTTP upload) is a
// first-class use of this library; the durable requirement is that commit RESOLVES so a
// caller can sequence upload finalization after it — never hangs waiting on a finish signal.
export async function streamCommitReport({duplex = false, timeoutMs = 4000} = {}) {
  const chunks = [];
  const stream = duplex
    ? new Duplex({read() {}, write(c, _e, cb) { chunks.push(c); cb(); }})
    : new PassThrough();
  if (!duplex) stream.on('data', c => chunks.push(c));

  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
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
      const back = new ExcelJS.Workbook();
      await back.xlsx.load(Buffer.concat(chunks));
      valid = back.worksheets.length === 1 && back.worksheets[0].getCell('A1').value === 'a';
    } catch {
      valid = false;
    }
  }
  return {settled, timedOut, bytes: Buffer.concat(chunks).length, valid};
}

// Report whether the streaming writer offers image parity with the in-memory writer: can a
// registered image be anchored onto a streamed worksheet, and does the streamed package then
// carry the media + drawing parts? A caller reaches for the streaming writer precisely for
// datasets too large to hold in memory, so a missing image path forces them back to the
// in-memory path they could not afford. Reports the capability surface and, if it works, the
// package facts — so a case can lock parity once the rewrite delivers it.
export async function streamWriterImageSupport(range = 'B2:D6') {
  const chunks = [];
  const stream = new PassThrough();
  stream.on('data', c => chunks.push(c));
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({stream});
  const sheet = writer.addWorksheet('S');
  const surface = {
    writerAddImage: typeof writer.addImage === 'function',
    sheetAddImage: typeof sheet.addImage === 'function',
  };
  let error = null;
  try {
    const imageId = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    sheet.addImage(imageId, range);
    sheet.addRow(['x']).commit();
    sheet.commit();
    await writer.commit();
  } catch (e) {
    error = String((e && e.message) || e);
  }

  let mediaParts = [];
  let drawingParts = [];
  if (!error && chunks.length) {
    const zip = await JSZip.loadAsync(Buffer.concat(chunks));
    mediaParts = Object.keys(zip.files).filter(n => /xl\/media\//.test(n));
    drawingParts = Object.keys(zip.files).filter(n => /drawing/.test(n));
  }
  return {...surface, error, mediaParts, drawingParts};
}
