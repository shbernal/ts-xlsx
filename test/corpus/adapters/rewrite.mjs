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

import {decodeAddress, decodeRange} from '../../../src/core/address.ts';
import {Workbook} from '../../../src/core/workbook.ts';

const impl = {
  name: 'rewrite',

  decodeAddress(reference) {
    return decodeAddress(reference);
  },

  decodeRange(reference) {
    return decodeRange(reference);
  },

  // Pure in-memory model capability: build a workbook, set a cell, and report the
  // runtime type and index of its position. Needs no serialization, so the core
  // model lights this up ahead of the reader/writer.
  cellColRowTypes(ref = 'B3') {
    const sheet = new Workbook().addWorksheet('S');
    const cell = sheet.getCell(ref);
    cell.value = 'x';
    return {col: cell.col, row: cell.row, colType: typeof cell.col, rowType: typeof cell.row};
  },
};

export default new Proxy(impl, {
  get(target, prop, receiver) {
    if (prop in target || typeof prop === 'symbol') {
      return Reflect.get(target, prop, receiver);
    }
    // A capability the rewrite has not reached yet. Return a thrower tagged so the
    // corpus runner distinguishes "not built here yet" (skip) from a real failure.
    return () => {
      const err = new Error(`rewrite: capability "${prop}" is not implemented yet`);
      err.notImplemented = true;
      throw err;
    };
  },
});
