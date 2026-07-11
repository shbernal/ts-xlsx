import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  MAX_COLUMN,
  columnToNumber,
  decodeAddress,
  decodeRange,
  encodeAddress,
  numberToColumn,
} from './address.ts';

test('numberToColumn covers the Excel range boundaries', () => {
  assert.equal(numberToColumn(1), 'A');
  assert.equal(numberToColumn(26), 'Z');
  assert.equal(numberToColumn(27), 'AA');
  assert.equal(numberToColumn(702), 'ZZ');
  assert.equal(numberToColumn(703), 'AAA');
  assert.equal(numberToColumn(MAX_COLUMN), 'XFD');
});

test('numberToColumn rejects out-of-bounds and non-integers', () => {
  assert.throws(() => numberToColumn(0), RangeError);
  assert.throws(() => numberToColumn(MAX_COLUMN + 1), RangeError);
  assert.throws(() => numberToColumn(1.5), RangeError);
});

test('columnToNumber is the inverse of numberToColumn across the whole range', () => {
  for (const n of [1, 26, 27, 52, 702, 703, 16383, MAX_COLUMN]) {
    assert.equal(columnToNumber(numberToColumn(n)), n);
  }
});

test('columnToNumber rejects invalid letters and overflow', () => {
  assert.throws(() => columnToNumber(''), RangeError);
  assert.throws(() => columnToNumber('a'), RangeError);
  assert.throws(() => columnToNumber('A1'), RangeError);
  assert.throws(() => columnToNumber('XFE'), RangeError); // 16385, just past XFD
});

test('decodeAddress reads a plain and an absolute cell identically', () => {
  assert.deepEqual(decodeAddress('B2'), {address: 'B2', col: 2, row: 2});
  assert.deepEqual(decodeAddress('$B$2'), {address: 'B2', col: 2, row: 2});
});

test('decodeAddress leaves the omitted axis undefined, not a sentinel', () => {
  assert.deepEqual(decodeAddress('$1'), {address: '1', col: undefined, row: 1});
  assert.deepEqual(decodeAddress('$A'), {address: 'A', col: 1, row: undefined});
});

test('decodeAddress rejects an empty reference', () => {
  assert.throws(() => decodeAddress('$'), SyntaxError);
  assert.throws(() => decodeAddress(''), SyntaxError);
});

test('decodeRange resolves an ordinary rectangle and normalizes corner order', () => {
  const range = decodeRange('B2:D6');
  assert.equal(range.top, 2);
  assert.equal(range.left, 2);
  assert.equal(range.bottom, 6);
  assert.equal(range.right, 4);
  assert.equal(range.dimensions, 'B2:D6');
  // reversed input yields the same normalized corners
  assert.deepEqual(decodeRange('D6:B2'), range);
});

test('decodeRange on a whole-row range leaks no undefined/NaN and keeps row bounds', () => {
  const range = decodeRange('$1:$1');
  assert.equal(range.top, 1);
  assert.equal(range.bottom, 1);
  assert.equal(range.left, undefined);
  assert.equal(range.right, undefined);
  assert.equal(range.dimensions, '1:1');
  const serialized = JSON.stringify(range);
  assert.ok(!serialized.includes('undefined'), serialized);
  assert.ok(!serialized.includes('NaN'), serialized);
});

test('decodeRange on a whole-column range is symmetric', () => {
  const range = decodeRange('$A:$C');
  assert.equal(range.left, 1);
  assert.equal(range.right, 3);
  assert.equal(range.top, undefined);
  assert.equal(range.bottom, undefined);
  assert.equal(range.dimensions, 'A:C');
});

test('decodeRange carries a quoted sheet name, unescaping doubled apostrophes', () => {
  assert.equal(decodeRange('Sheet1!A1:B2').sheetName, 'Sheet1');
  assert.equal(decodeRange("'Bob''s data'!A1:B2").sheetName, "Bob's data");
  assert.equal(decodeRange('A1:B2').sheetName, undefined);
});

test('encodeAddress round-trips with decodeAddress', () => {
  assert.equal(encodeAddress(2, 6), 'B6');
  assert.equal(encodeAddress(MAX_COLUMN, 1048576), 'XFD1048576');
  const decoded = decodeAddress(encodeAddress(30, 42));
  assert.equal(decoded.col, 30);
  assert.equal(decoded.row, 42);
});
