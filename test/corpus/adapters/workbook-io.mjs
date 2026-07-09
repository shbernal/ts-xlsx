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
//     sheets: [{
//       name,                                                       // required
//       cells:   [{ ref, value|formula(+result)|text+hyperlink, numFmt, font }],
//       columns: [{ index, width, hidden }],                        // index: 1-based
//       rows:    [{ index, height, hidden }],
//       pageMargins: { left, right, top, bottom, header, footer },  // any subset
//       tables:  [{ name, ref, headers:[…], rows:[[…]], totalsRow }],
//     }],
//   }
// A cell value of { invalidDate: true } materializes `new Date(NaN)`.

import {createRequire} from 'node:module';

const require = createRequire(import.meta.url);
const ExcelJS = require('../../../lib/exceljs.nodejs.js');
const JSZip = require('jszip');

const toDate = v => (v && typeof v === 'object' && v.invalidDate ? new Date(NaN) : new Date(v));

function cellValueFrom(c) {
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
    }
    for (const col of s.columns || []) {
      const column = sheet.getColumn(col.index);
      if (col.width !== undefined) column.width = col.width;
      if (col.hidden !== undefined) column.hidden = col.hidden;
    }
    for (const row of s.rows || []) {
      const r = sheet.getRow(row.index);
      if (row.height !== undefined) r.height = row.height;
      if (row.hidden !== undefined) r.hidden = row.hidden;
    }
    // Assign only the margins the spec provides — faithfully reproducing the user
    // scenario of setting a subset (do NOT pre-fill defaults; that is exactly the
    // write-side behavior under test).
    if (s.pageMargins) sheet.pageSetup.margins = {...s.pageMargins};
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
  }
  return workbook;
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
      columns[col.index] = {width: column.width ?? null, hidden: !!column.hidden};
    }
    const rows = {};
    for (const row of s.rows || []) {
      const r = sheet.getRow(row.index);
      rows[row.index] = {height: r.height ?? null, hidden: !!r.hidden};
    }
    sheets[s.name] = {
      cells,
      columns,
      rows,
      margins: sheet.pageSetup ? {...sheet.pageSetup.margins} : null,
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
    sheets,
  };
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
  const overrides = [...contentTypes.matchAll(/<Override[^>]*PartName="([^"]*)"[^>]*\/>/g)].map(m => m[1]);
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
    sheets[s.name] = {
      pageMargins: {present: Object.keys(marginAttrs), values: marginAttrs},
      hasSheetViews: /<sheetViews>/.test(xml),
      sheetViewCount: sheetViewTags.length,
      hasDimension: /<dimension\b/.test(xml),
      formulas,
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
    sheetEntries,
    rels,
    sheets,
    tables,
    consistency: {
      worksheetPartCount: worksheetParts.length,
      sheetEntryCount: sheetEntries.length,
      declaredConsistent,
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

function xmlWellFormed(xml) {
  // Cheap structural check: no raw & that isn't an entity, tags balanced enough to
  // matter. A real parser lives in the impl; here we only need "would a strict
  // consumer choke", which unescaped & or < inside text triggers.
  if (/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml)) return false;
  return true;
}
