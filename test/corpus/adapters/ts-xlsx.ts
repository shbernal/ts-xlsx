// Adapter: binds the corpus's contract vocabulary to this library's public API.
//
// Node runs the .ts sources directly (type-stripping), so this imports them with no build step; see
// ./ts-xlsx/runtime.ts for the src-vs-dist switch that decides *which* copy is under test.
//
// The capabilities live in ./ts-xlsx/, one module per concern; this file only assembles them. Nothing
// here closes over anything, so a capability can move between modules without any case noticing.
//
// There used to be a Proxy here that turned an unknown property into a thrower tagged for the runner
// to report as "skipped" — the right design while the library was being rebuilt module by module and a
// case could legitimately arrive before the code it tested. That is over, and the same mechanism now
// only has one possible cause: a case calling a capability that does not exist. Skipping that silently
// is the worst available outcome, because the case reads as accounted for. Deleting the Proxy makes it
// the best one — with `CorpusApi` derived from this object, the typo is a compile error naming the
// capability, caught by `typecheck:test` before anything runs.

import {comments} from './ts-xlsx/comments.ts';
import {conditionalFormatting} from './ts-xlsx/conditional-formatting.ts';
import {core} from './ts-xlsx/core.ts';
import {csv} from './ts-xlsx/csv.ts';
import {formulas} from './ts-xlsx/formulas.ts';
import {grid} from './ts-xlsx/grid.ts';
import {hyperlinks} from './ts-xlsx/hyperlinks.ts';
import {images} from './ts-xlsx/images.ts';
import {protection} from './ts-xlsx/protection.ts';
import {streaming} from './ts-xlsx/streaming.ts';
import {styles} from './ts-xlsx/styles.ts';
import {tables} from './ts-xlsx/tables.ts';
import {validation} from './ts-xlsx/validation.ts';
import {vba} from './ts-xlsx/vba.ts';
import {xlsb} from './ts-xlsx/xlsb.ts';

const impl = {
  name: 'ts-xlsx',
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
  ...xlsb,
};

/**
 * The capability surface a case calls, derived from the assembled adapter rather than declared
 * alongside it.
 *
 * Derived, because a hand-written contract interface would be a second list of ~180 signatures to keep
 * in step with these fifteen modules, and this project's answer to a second list is always to generate
 * it from the first (`WORKSHEET_MODEL_FACETS`, `check-entries.ts`). Derivation makes drift impossible
 * instead of merely detectable.
 *
 * This is what the cases are blind to and what they are not. They still cannot see `src` internals —
 * they reach the library only through the capabilities below, which is the decoupling that let the
 * corpus outlive the rewrite. What they no longer pretend not to see is the *shape of this surface*:
 * a capability's name, its arguments and what it returns are the contract a case is written against,
 * so hiding them behind `any` bought no independence and cost every case its type-checking.
 */
export type CorpusApi = typeof impl;

export default impl;
