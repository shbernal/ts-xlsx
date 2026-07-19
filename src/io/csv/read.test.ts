import assert from 'node:assert/strict';
import test from 'node:test';

import {readCsv} from './read.ts';

// The single worksheet's rows as a plain 2-D array of cell values, for terse assertions.
function rowsOf(csv: string, options?: Parameters<typeof readCsv>[1]): unknown[][] {
  const sheet = readCsv(csv, options).worksheets[0]!;
  const rows: unknown[][] = [];
  for (const {cells} of sheet.rows()) {
    let width = 0;
    for (const cell of cells) if (cell.col > width) width = cell.col;
    const fields: unknown[] = new Array(width).fill(null);
    for (const cell of cells) fields[cell.col - 1] = cell.value;
    rows.push(fields);
  }
  return rows;
}

test('a configured delimiter splits fields and numeric fields coerce', () => {
  assert.deepEqual(rowsOf('a;b;c\n1;2;3', {delimiter: ';'}), [
    ['a', 'b', 'c'],
    [1, 2, 3],
  ]);
});

test('an over-precision numeric string is preserved verbatim; in-range numbers coerce', () => {
  const big = '56343416020533614003';
  assert.deepEqual(rowsOf(`${big},42\n1.5,7`), [
    [big, 42],
    [1.5, 7],
  ]);
});

test('a leading-zero id coerces to a number by default but survives the identity map', () => {
  assert.deepEqual(rowsOf('007,32.5'), [[7, 32.5]]);
  assert.deepEqual(rowsOf('007,32.5', {map: (v) => v}), [['007', '32.5']]);
});

test('padded ids and dash-codes stay strings; a strict ISO date becomes a Date', () => {
  assert.deepEqual(rowsOf('2020-00001,1-3,3-4'), [['2020-00001', '1-3', '3-4']]);
  const cell = rowsOf('2018-01-05')[0]![0];
  assert.ok(cell instanceof Date);
  assert.equal(cell.toISOString(), '2018-01-05T00:00:00.000Z');
});

test('a whitespace-only field is a string, an empty field is null — neither is 0', () => {
  const [row] = rowsOf('firstValue,   ,secondValue\n');
  assert.equal(typeof row![1], 'string');
  assert.notEqual(row![1], 0);
  assert.equal(rowsOf('firstValue,,secondValue')[0]![1], null);
});

test('header mode consumes the first line, leaving data rows', () => {
  assert.deepEqual(rowsOf('name,age\nalice,30', {headers: true}), [['alice', 30]]);
  assert.deepEqual(rowsOf('name,age\nalice,30'), [
    ['name', 'age'],
    ['alice', 30],
  ]);
});

test('quoted fields carry embedded delimiters, quotes, and newlines', () => {
  assert.deepEqual(rowsOf('"a,b","he said ""hi""","line\nbreak"'), [
    ['a,b', 'he said "hi"', 'line\nbreak'],
  ]);
});

test('a leading UTF-8 BOM is stripped and bytes decode', () => {
  const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...Buffer.from('x,y', 'utf8')]);
  assert.deepEqual(rowsOfBytes(withBom), [['x', 'y']]);
});

function rowsOfBytes(bytes: Uint8Array): unknown[][] {
  const sheet = readCsv(bytes).worksheets[0]!;
  const rows: unknown[][] = [];
  for (const {cells} of sheet.rows()) rows.push(cells.map((c) => c.value));
  return rows;
}
