import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../src/index.ts';
import {difference, read, roundTrip, write} from './lanes.ts';
import {findSample} from './samples.ts';

function sampleBytes(id: string): Uint8Array {
  const sample = findSample(id);
  assert.ok(sample !== undefined, id);
  const written = write(sample.build());
  assert.ok(written.ok);
  return written.value.bytes;
}

test('a workbook with no worksheets fails the write lane as a result, not a throw', () => {
  // The page has to render this: a thrown error there is a blank panel, and a reader who
  // learns the library refuses a zero-sheet package by name has learned something true.
  const result = write(new Workbook());
  assert.equal(result.ok, false);
  assert.ok(result.ok || result.error.length > 0);
});

test('bytes that are not a package fail the read lane by name', () => {
  const result = read(new Uint8Array([1, 2, 3, 4]));
  assert.equal(result.ok, false);
  assert.match(result.ok ? '' : result.error, /Error/);
});

test('the read lane reports a sheet and the range it occupies', () => {
  const result = read(sampleBytes('structure'));
  assert.ok(result.ok);
  const [sheet] = result.value.sheets;
  assert.ok(sheet !== undefined);
  assert.equal(sheet.name, 'Structure');
  assert.equal(sheet.usedRange, 'A1:C6');
  assert.ok(sheet.rowCount > 0);
});

test('a workbook this library wrote itself has nothing it had to preserve verbatim', () => {
  const result = read(sampleBytes('values'));
  assert.ok(result.ok);
  assert.deepEqual(result.value.preservedParts, []);
});

test('every lane reports a finite, non-negative elapsed time', () => {
  const bytes = sampleBytes('values');
  const readBack = read(bytes);
  assert.ok(readBack.ok);
  const trip = roundTrip(bytes);
  assert.ok(trip.ok);
  for (const ms of [readBack.value.elapsedMs, trip.value.elapsedMs]) {
    assert.ok(Number.isFinite(ms) && ms >= 0, String(ms));
  }
});

test('difference names the path where two models disagree', () => {
  assert.equal(difference({a: 1}, {a: 1}, ''), undefined);
  assert.equal(difference({a: 1}, {a: 2}, ''), 'a');
  assert.equal(difference({a: {b: 'x'}}, {a: {b: 'y'}}, ''), 'a.b');
  assert.equal(difference([{n: 1}, {n: 2}], [{n: 1}, {n: 3}], ''), '[1].n');
});

test('difference treats a missing key and a null the same, and a length change as one', () => {
  // Two reads of the same file must not disagree over whether an unset facet is absent or
  // null; those are the same fact written two ways, and reporting it would be a false alarm.
  assert.equal(difference({a: 1, b: null}, {a: 1}, ''), undefined);
  assert.equal(difference([1, 2], [1, 2, 3], ''), '(root).length');
});

test('difference reports the root when the two are not even the same shape', () => {
  assert.equal(difference([1], {0: 1}, ''), '(root)');
  assert.equal(difference('x', 1, ''), '(root)');
});
