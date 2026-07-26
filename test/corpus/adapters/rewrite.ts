// Adapter: binds the corpus's implementation-blind contract vocabulary to the
// *rewrite* — the new, strict-TypeScript library under src/ (Phase 3). This is the
// sibling of current.mjs: the same cases run unchanged against it, and the finish
// line for an area is every one of its baselines flipping to `pass` here.
//
// Node 24 runs the .ts sources directly (type-stripping), so this adapter imports
// them with no build step. Strict type-checking is enforced separately by
// `pnpm run typecheck` (tsc --noEmit -p tsconfig.json).
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
//
// The capabilities themselves live in ./rewrite/, one module per concern; this file is the
// assembly point and the feature gate. Nothing here closes over anything, so a capability
// can move between modules without any case noticing.

import {comments} from './rewrite/comments.ts';
import {conditionalFormatting} from './rewrite/conditional-formatting.ts';
import {core} from './rewrite/core.ts';
import {csv} from './rewrite/csv.ts';
import {formulas} from './rewrite/formulas.ts';
import {grid} from './rewrite/grid.ts';
import {hyperlinks} from './rewrite/hyperlinks.ts';
import {images} from './rewrite/images.ts';
import {protection} from './rewrite/protection.ts';
import {notImplemented} from './rewrite/spec-model.ts';
import {streaming} from './rewrite/streaming.ts';
import {styles} from './rewrite/styles.ts';
import {tables} from './rewrite/tables.ts';
import {validation} from './rewrite/validation.ts';
import {vba} from './rewrite/vba.ts';

const impl = {
  name: 'rewrite',
  ...comments,
  ...conditionalFormatting,
  ...core,
  ...csv,
  ...formulas,
  ...grid,
  ...hyperlinks,
  ...images,
  ...protection,
  ...streaming,
  ...styles,
  ...tables,
  ...validation,
  ...vba,
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
