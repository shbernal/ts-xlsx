// Smoke test for the publishable build.
//
// Typecheck proves the source is sound; it cannot prove the *emitted* artifact loads and runs.
// This imports the package exactly as a consumer would and exercises a write → read round-trip.
// It guards against emit-shaped regressions: a broken import specifier rewrite, a missing file, a
// runtime-only reference that type-stripping tolerated.
//
// Imports go through the package NAME, not a relative `../dist/` path, so Node resolves them the
// way a consumer's would, through `package.json`'s `exports` map (self-reference, which Node
// enables for any package that declares `exports`). That is the only thing in the repo that
// exercises the map: the corpus's dist target loads emitted modules by file path, so a subpath
// that resolved to nothing would pass every other gate and fail on a consumer's first install.
//
// Each subpath is checked for a binding that must be there, and the module graph is checked for
// one that must NOT be: `/core` reaching a codec, or `/errors` reaching anything at all, is the
// packaging regression these entry points exist to prevent.
//
// Findings accumulate and are reported through `verdict` at the end, like every gate in this
// directory: the boundary checks are the ones where the whole list is the diagnostic. The browser
// boundary is not re-checked here. `scripts/check-browser-safe.ts` proves it over src/, and the
// emit (`verbatimModuleSyntax`, no helpers) can only rewrite or drop an import, never add one.

import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {isDeepStrictEqual} from 'node:util';

import {closure, importedPaths} from './module-graph.ts';
import {verdict} from './verdict.ts';

const problems: string[] = [];
const check = (held: boolean, problem: string): void => {
  if (!held) problems.push(`  ${problem}`);
};

const {decodeAddress, readXlsx, Workbook, writeXlsx} = await import('@shbernal/ts-xlsx');

const wb = new Workbook();
const ws = wb.addWorksheet('Smoke');
ws.getCell('A1').value = 'hello';
ws.getCell('B2').value = 42;

const bytes = writeXlsx(wb);
const roundTrip = readXlsx(bytes);
const sheet = roundTrip.getWorksheet('Smoke');
check(sheet !== undefined, 'round-tripped workbook lost the worksheet');
check(sheet?.getCell('A1').value === 'hello', 'A1 did not survive round-trip');
check(sheet?.getCell('B2').value === 42, 'B2 did not survive round-trip');

check(
  isDeepStrictEqual(decodeAddress('B2'), {address: 'B2', col: 2, row: 2}),
  'address decode wrong',
);

// One binding per subpath: enough to prove the specifier resolves to the right module, not a
// re-listing of every export (`scripts/check-entries.ts` owns that).
const SUBPATH_BINDINGS: Readonly<Record<string, string>> = {
  core: 'Workbook',
  xlsx: 'readXlsx',
  xlsb: 'readXlsb',
  csv: 'readCsv',
  node: 'WorkbookStreamWriter',
  vba: 'parseVbaProject',
  customui: 'parseCustomUi',
  errors: 'XlsxError',
};

for (const [subpath, binding] of Object.entries(SUBPATH_BINDINGS)) {
  const specifier = `@shbernal/ts-xlsx/${subpath}`;
  const module = (await import(specifier)) as Record<string, unknown>;
  check(binding in module, `${specifier} does not export ${binding}`);
}

// What each entry point actually costs, checked as a shape rather than a size: `scripts/
// size-budget.ts` catches growth, this catches a boundary being crossed at all.
const HERE = dirname(fileURLToPath(import.meta.url));
const entryFile = (subpath: string) => join(HERE, '..', 'dist', 'entries', `${subpath}.js`);

const coreReach = closure(entryFile('core'), importedPaths);
for (const file of coreReach) {
  check(
    !file.includes(join('dist', 'io')),
    `/core reaches ${file}: the model must not pull in a serialisation`,
  );
}

const errorsReach = closure(entryFile('errors'), importedPaths);
// The taxonomy, and the leaves the taxonomy is built from. `hex.js` is the one leaf today: an error
// naming a code point renders it as hex, and `src/hex.ts` is its own module precisely so that this
// entry can have it without taking `bytes.ts` with it. The rule is that nothing here parses anything,
// which is why the allowance is a named list rather than a relaxed predicate.
const ERRORS_LEAVES = new Set(['hex.js']);
for (const file of errorsReach) {
  const name = file.slice(file.lastIndexOf('/') + 1);
  check(
    file.endsWith('errors.js') || ERRORS_LEAVES.has(name),
    `/errors reaches ${file}: the taxonomy must cost nothing but itself and its named leaves`,
  );
}

verdict({
  gate: 'dist smoke',
  problems,
  ok:
    `${bytes.byteLength} byte xlsx, round-trip verified; ` +
    `${Object.keys(SUBPATH_BINDINGS).length} subpaths resolve through exports; ` +
    `/core is codec-free (${coreReach.size} modules), /errors is self-contained (${errorsReach.size})`,
  failure: 'problem(s) in the published build',
});
