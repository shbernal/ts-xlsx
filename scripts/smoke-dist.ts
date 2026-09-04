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
// directory. Written as bare `assert` calls, this stopped at the first: a build that had put three
// Node built-ins on the browser path reported one, and the next run reported the next. The boundary
// checks are exactly the ones where the whole list is the diagnostic.

import {readFileSync} from 'node:fs';
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
check(bytes.byteLength > 0, 'writer produced no bytes');
check(bytes[0] === 0x50 && bytes[1] === 0x4b, 'output is not a zip (bad PK magic)');

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
// re-listing of the export surface (`scripts/check-entries.ts` owns that).
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

// The browser boundary, checked on the EMITTED artifact rather than on source.
// `scripts/check-browser-safe.ts` proves it over src/ before anything is built; this proves the
// emit kept it, which is the only form a consumer's bundler ever sees. `/node` is the exception
// that gives the rule its shape: it is where those imports are allowed to be.
const NODE_SPECIFIER = /\b(?:from|import)\s+["']node:/;
const browserEntries = Object.keys(SUBPATH_BINDINGS).filter((subpath) => subpath !== 'node');
for (const subpath of [...browserEntries, 'index']) {
  const entry = subpath === 'index' ? join(HERE, '..', 'dist', 'index.js') : entryFile(subpath);
  for (const file of closure(entry, importedPaths)) {
    check(
      !NODE_SPECIFIER.test(readFileSync(file, 'utf8')),
      `/${subpath} reaches ${file}, which imports a Node built-in: a browser cannot bundle it`,
    );
  }
}
// And from the other side: `/node` must still reach one, or the boundary has moved rather than
// held and the browser assertions above are passing for the wrong reason.
const nodeClosure = [...closure(entryFile('node'), importedPaths)];
check(
  nodeClosure.some((file) => NODE_SPECIFIER.test(readFileSync(file, 'utf8'))),
  '/node reaches no Node built-in: the streaming writer it exists to carry is not behind it',
);

verdict({
  gate: 'dist smoke',
  problems,
  ok:
    `${bytes.byteLength} byte xlsx, round-trip verified; ` +
    `${Object.keys(SUBPATH_BINDINGS).length} subpaths resolve through exports; ` +
    `/core is codec-free (${coreReach.size} modules), /errors is self-contained (${errorsReach.size}); ` +
    `${browserEntries.length + 1} browser entries import no Node built-in`,
  failure: 'problem(s) in the published build',
});
