import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from './workbook.ts';
import type {Worksheet} from './worksheet.ts';

function sheet(): Worksheet {
  return new Workbook().addWorksheet('S');
}

test('a range is reachable by reference and by inclusive corners, and the two agree', () => {
  const s = sheet();
  const byRef = s.getRange('B2:D5');
  const byCorners = s.getRange(2, 2, 5, 4);
  assert.equal(byRef.address, 'B2:D5');
  assert.equal(byCorners.address, 'B2:D5');
  assert.deepEqual(
    {top: byRef.top, left: byRef.left, bottom: byRef.bottom, right: byRef.right},
    {top: 2, left: 2, bottom: 5, right: 4},
  );
});

test('corners are normalised, so a reversed reference names the same block', () => {
  const s = sheet();
  assert.equal(s.getRange('D5:B2').address, 'B2:D5');
  assert.equal(s.getRange(5, 4, 2, 2).address, 'B2:D5');
});

test('a single-cell reference is a degenerate one-cell range', () => {
  const range = sheet().getRange('C3');
  assert.equal(range.address, 'C3:C3');
  assert.equal(range.cellCount, 1);
});

test('a range reports its geometry without touching the grid', () => {
  const s = sheet();
  const range = s.getRange('B2:D5');
  assert.equal(range.rowCount, 4);
  assert.equal(range.columnCount, 3);
  assert.equal(range.cellCount, 12);
  // Asking about a block must cost nothing — no cells, and no growth of the used range.
  assert.equal(s.rowCount, 0);
  assert.equal(s.hasCell(2, 2), false);
});

test('addresses walk the block row-major and materialise nothing', () => {
  const s = sheet();
  const range = s.getRange('B2:C3');
  assert.deepEqual([...range.addresses()], ['B2', 'C2', 'B3', 'C3']);
  assert.equal(s.hasCell(2, 2), false, 'walking addresses creates no cells');
});

test('addresses is a generator, so a large block can be abandoned part-way', () => {
  // The point of not returning an array: `A1:XFD1048576` is 17 billion addresses.
  const range = sheet().getRange(1, 1, 1048576, 16384);
  const first: string[] = [];
  for (const address of range.addresses()) {
    first.push(address);
    if (first.length === 3) break;
  }
  assert.deepEqual(first, ['A1', 'B1', 'C1']);
});

test('cells reports only what exists, sparsely, and creates nothing', () => {
  const s = sheet();
  s.getCell('C2').value = 'x';
  s.getCell('B3').value = 'y';
  const range = s.getRange('B2:C3');
  assert.deepEqual(
    range.cells.map((cell) => cell.address),
    ['C2', 'B3'],
    'row-major over the materialised cells only',
  );
  assert.equal(s.hasCell(2, 2), false, 'reading the block did not fill the holes');
});

test('contains answers for positions inside and outside the block', () => {
  const range = sheet().getRange('B2:D5');
  assert.equal(range.contains(2, 2), true);
  assert.equal(range.contains(5, 4), true, 'both edges are inclusive');
  assert.equal(range.contains(1, 2), false);
  assert.equal(range.contains(2, 5), false);
});

test('a whole-row or whole-column reference is refused, pointing at the axis handle', () => {
  // Not a size limit but a shape mismatch: OOXML says "this whole column is bold" in one attribute,
  // and accepting `A:A` here would silently mean a million cell objects saying it one at a time.
  const s = sheet();
  assert.throws(() => s.getRange('A:A'), /whole columns.*getColumn/s);
  assert.throws(() => s.getRange('1:1'), /whole rows.*getRow/s);
  assert.throws(() => s.getRange('A:C'), /whole columns/);
});

test('a reference naming another worksheet is refused rather than silently retargeted', () => {
  const wb = new Workbook();
  const a = wb.addWorksheet('Alpha');
  wb.addWorksheet('Beta');
  assert.throws(() => a.getRange('Beta!B2:D5'), /names worksheet "Beta", not "Alpha"/);
  // Its own name is fine, and matches case-insensitively as sheet lookup does everywhere else.
  assert.equal(a.getRange('alpha!B2:D5').address, 'B2:D5');
});

test('an out-of-bounds or non-integer corner is refused', () => {
  const s = sheet();
  assert.throws(() => s.getRange(0, 1, 2, 2), /row 0 is out of bounds/);
  assert.throws(() => s.getRange(1, 1, 2, 16385), /column 16385 is out of bounds/);
  assert.throws(() => s.getRange(1, 1, 1048577, 2), /row 1048577 is out of bounds/);
  assert.throws(() => s.getRange(1.5, 1, 2, 2), /row 1.5 is out of bounds/);
});

test('a range knows the sheet it came from', () => {
  const s = sheet();
  assert.equal(s.getRange('A1:B2').sheet, s);
});
