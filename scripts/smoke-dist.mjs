// Smoke test for the publishable build.
//
// Typecheck proves the source is sound; it cannot prove the *emitted* artifact
// loads and runs. This imports the package exactly as a consumer would — through
// the compiled dist/index.js — and exercises a write → read round-trip. It guards
// against emit-shaped regressions: a broken import specifier rewrite, a missing
// file, a runtime-only reference to something type-stripping tolerated.

import assert from 'node:assert/strict';
import {Workbook, writeXlsx, readXlsx, decodeAddress} from '../dist/index.js';

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

console.log(`dist smoke ok — ${bytes.byteLength} byte xlsx, round-trip verified`);
