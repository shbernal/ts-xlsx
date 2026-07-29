// Smoke test for the publishable build.
//
// Typecheck proves the source is sound; it cannot prove the *emitted* artifact loads and runs.
// This imports the package exactly as a consumer would and exercises a write → read round-trip.
// It guards against emit-shaped regressions: a broken import specifier rewrite, a missing file, a
// runtime-only reference that type-stripping tolerated.
//
// Imports go through the package NAME, not a relative `../dist/` path, so Node resolves them the
// way a consumer's would — through `package.json`'s `exports` map (self-reference, which Node
// enables for any package that declares `exports`). That is the only thing in the repo that
// exercises the map: the corpus's dist target loads emitted modules by file path, so a subpath
// that resolved to nothing would pass every other gate and fail on a consumer's first install.
//
// Each subpath is checked for a binding that must be there, and the module graph is checked for
// one that must NOT be: `/core` reaching a codec, or `/errors` reaching anything at all, is the
// packaging regression these entry points exist to prevent.

import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {dirname, join, normalize} from 'node:path';
import {fileURLToPath} from 'node:url';

const {decodeAddress, readXlsx, Workbook, writeXlsx} = await import('@shbernal/ts-xlsx');

const wb = new Workbook();
const ws = wb.addWorksheet('Smoke');
ws.getCell('A1').value = 'hello';
ws.getCell('B2').value = 42;

const bytes = await writeXlsx(wb);
assert.ok(bytes.byteLength > 0, 'writer produced no bytes');
assert.ok(bytes[0] === 0x50 && bytes[1] === 0x4b, 'output is not a zip (bad PK magic)');

const roundTrip = readXlsx(bytes);
const sheet = roundTrip.getWorksheet('Smoke');
assert.ok(sheet, 'round-tripped workbook lost the worksheet');
assert.equal(sheet.getCell('A1').value, 'hello', 'A1 did not survive round-trip');
assert.equal(sheet.getCell('B2').value, 42, 'B2 did not survive round-trip');

assert.deepEqual(decodeAddress('B2'), {address: 'B2', col: 2, row: 2}, 'address decode wrong');

// One binding per subpath: enough to prove the specifier resolves to the right module, not a
// re-listing of the export surface (`scripts/check-entries.ts` owns that).
const SUBPATH_BINDINGS: Readonly<Record<string, string>> = {
  core: 'Workbook',
  xlsx: 'readXlsx',
  xlsb: 'readXlsb',
  csv: 'readCsv',
  vba: 'parseVbaProject',
  customui: 'parseCustomUi',
  errors: 'XlsxError',
};

for (const [subpath, binding] of Object.entries(SUBPATH_BINDINGS)) {
  const specifier = `@shbernal/ts-xlsx/${subpath}`;
  const module = (await import(specifier)) as Record<string, unknown>;
  assert.ok(binding in module, `${specifier} does not export ${binding}`);
}

// What each entry point actually costs, checked as a shape rather than a size: `scripts/
// size-budget.ts` catches growth, this catches a boundary being crossed at all.
const HERE = dirname(fileURLToPath(import.meta.url));
const RELATIVE_SPECIFIER = /\bfrom\s+["'](\.[^"']*)["']/g;

function closure(entry: string): Set<string> {
  const reached = new Set<string>();
  const pending = [entry];
  while (pending.length > 0) {
    const file = pending.pop() as string;
    if (reached.has(file)) continue;
    reached.add(file);
    const source = readFileSync(file, 'utf8');
    for (const match of source.matchAll(RELATIVE_SPECIFIER)) {
      pending.push(normalize(join(dirname(file), match[1] as string)));
    }
  }
  return reached;
}

const entryFile = (subpath: string) => join(HERE, '..', 'dist', 'entries', `${subpath}.js`);

const coreReach = closure(entryFile('core'));
for (const file of coreReach) {
  assert.ok(
    !file.includes(join('dist', 'io')),
    `/core reaches ${file} — the model must not pull in a serialisation`,
  );
}

const errorsReach = closure(entryFile('errors'));
for (const file of errorsReach) {
  assert.ok(
    file.endsWith('errors.js'),
    `/errors reaches ${file} — the taxonomy must cost nothing but itself`,
  );
}

console.log(
  `dist smoke ok — ${bytes.byteLength} byte xlsx, round-trip verified; ` +
    `${Object.keys(SUBPATH_BINDINGS).length} subpaths resolve through exports; ` +
    `/core is codec-free (${coreReach.size} modules), /errors is self-contained (${errorsReach.size})`,
);
