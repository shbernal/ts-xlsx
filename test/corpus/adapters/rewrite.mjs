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
import {writeXlsx} from '../../../src/io/xlsx/write.ts';
import {packageFacts} from './ooxml-facts.mjs';

const notImplemented = message => {
  const err = new Error(`rewrite: ${message}`);
  err.notImplemented = true;
  return err;
};

const SUPPORTED_TOP_KEYS = new Set(['sheets', 'properties']);
const SUPPORTED_PROP_KEYS = new Set(['creator', 'lastModifiedBy', 'created', 'modified']);
const SUPPORTED_SHEET_KEYS = new Set(['name', 'state', 'cells']);
const SUPPORTED_CELL_KEYS = new Set(['ref', 'value', 'formula', 'result']);

const toDate = v => (v && typeof v === 'object' && v.invalidDate ? new Date(NaN) : new Date(v));

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
    for (const c of s.cells || []) {
      for (const k of Object.keys(c)) {
        if (!SUPPORTED_CELL_KEYS.has(k)) throw notImplemented(`cell.${k} not supported yet`);
      }
      const cell = sheet.getCell(c.ref);
      if ('formula' in c) {
        cell.value = 'result' in c ? {formula: c.formula, result: c.result} : {formula: c.formula};
      } else {
        const v = c.value;
        if (v !== null && typeof v === 'object') {
          throw notImplemented(`cell value shape ${JSON.stringify(v)} not supported yet`);
        }
        cell.value = v;
      }
    }
  }
  return workbook;
}

function partMapOf(buffer) {
  const unzipped = unzipSync(buffer);
  const out = {};
  for (const name of Object.keys(unzipped)) out[name] = strFromU8(unzipped[name]);
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
    // The write succeeded, but reporting which cells survived requires reading the
    // package back — the reader slice will provide that, so skip until then.
    throw notImplemented('tryWriteWorkbook read-back needs the reader');
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
