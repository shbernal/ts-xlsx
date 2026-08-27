import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {dirname, resolve} from 'node:path';
import {test} from 'node:test';
import {fileURLToPath} from 'node:url';

import {read, roundTrip, write} from './lanes.ts';
import {builderSource, findSample, SAMPLES} from './samples.ts';

const MODULE_TEXT = readFileSync(
  resolve(dirname(fileURLToPath(import.meta.url)), 'samples.ts'),
  'utf8',
);

test('every sample has a distinct id', () => {
  const ids = SAMPLES.map((sample) => sample.id);
  assert.equal(new Set(ids).size, ids.length);
});

test('every sample builds, writes and reads back', () => {
  for (const sample of SAMPLES) {
    const written = write(sample.build());
    assert.ok(written.ok, `${sample.id} failed to write: ${written.ok ? '' : written.error}`);
    assert.ok(written.value.byteLength > 0, `${sample.id} wrote an empty package`);

    const readBack = read(written.value.bytes);
    assert.ok(readBack.ok, `${sample.id} failed to read: ${readBack.ok ? '' : readBack.error}`);
    assert.ok(readBack.value.sheets.length > 0, `${sample.id} read back with no sheets`);
  }
});

test('every sample round-trips in agreement', () => {
  // A failure here is a library bug, and a good one to have found: it means a model written
  // and read back does not match the one that produced it.
  for (const sample of SAMPLES) {
    const written = write(sample.build());
    assert.ok(written.ok);
    const trip = roundTrip(written.value.bytes);
    assert.ok(trip.ok, `${sample.id} failed to round-trip: ${trip.ok ? '' : trip.error}`);
    assert.equal(
      trip.value.firstDisagreement,
      undefined,
      `${sample.id} disagreed at ${trip.value.firstDisagreement ?? ''}`,
    );
    assert.ok(trip.value.bytesIdentical, `${sample.id} wrote two different packages`);
  }
});

test('the values sample carries back every kind it was given', () => {
  const sample = findSample('values');
  assert.ok(sample !== undefined);
  const written = write(sample.build());
  assert.ok(written.ok);
  const readBack = read(written.value.bytes);
  assert.ok(readBack.ok);

  const sheet = readBack.value.workbook.requireWorksheet('Values');
  assert.equal(sheet.getCell('B2').value, 1234.5);
  assert.equal(sheet.getCell('B3').value, 'hello');
  assert.equal(sheet.getCell('B4').value, true);
  assert.ok(sheet.getCell('B5').value instanceof Date);
  assert.equal(sheet.getCell('B5').numFmt, 'yyyy-mm-dd');
  assert.deepEqual(sheet.getCell('B7').value, {error: '#DIV/0!'});
});

test('the corpus sample carries its format code back character for character', () => {
  // The same assertion `custom-numfmt-string-roundtrips-verbatim` makes, so that the claim
  // the page puts on screen is the claim the corpus already holds the library to.
  const sample = findSample('corpus-numfmt');
  assert.ok(sample !== undefined);
  const declared = sample.build().requireWorksheet('NumFmt').getCell('B2').numFmt;
  const written = write(sample.build());
  assert.ok(written.ok);
  const readBack = read(written.value.bytes);
  assert.ok(readBack.ok);
  assert.equal(readBack.value.workbook.requireWorksheet('NumFmt').getCell('B2').numFmt, declared);
});

test('every builder can be cut out of this module for the page to show', () => {
  for (const sample of SAMPLES) {
    const source = builderSource(MODULE_TEXT, sample.id);
    assert.ok(source.startsWith(`function ${sample.build.name}(`), sample.id);
    assert.ok(source.endsWith('}'), sample.id);
    assert.ok(source.includes('return workbook;'), sample.id);
  }
});

test('asking for a builder that is not there says so rather than showing nothing', () => {
  assert.throws(() => builderSource(MODULE_TEXT, 'no-such-sample'), /no sample with id/);
  assert.throws(() => builderSource('', 'values'), /not a top-level function declaration/);
});
