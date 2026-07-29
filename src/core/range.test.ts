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

// --- styling the block ---------------------------------------------------------------------------

test('assigning a facet applies it to every cell, materialising the empty ones', () => {
  // A styled-but-valueless cell is the only way an empty cell renders with a fill; skipping the
  // holes would leave gaps in a header band, which is the whole motivating case.
  const s = sheet();
  s.getCell('B2').value = 'header';
  const range = s.getRange('B2:C3');
  range.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFEEEEEE'}};

  assert.equal(range.cells.length, 4, 'every position in the block now exists');
  for (const cell of range.cells) {
    assert.deepEqual(cell.fill, {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFEEEEEE'}});
  }
  assert.equal(s.getCell('C3').value, null, 'a materialised cell carries style, not a value');
  assert.equal(s.getCell('D2').fill, undefined, 'nothing outside the block was touched');
});

test('a facet reads back only when every cell in the block agrees', () => {
  const s = sheet();
  const range = s.getRange('B2:C3');
  range.numFmt = '0.00';
  assert.equal(range.numFmt, '0.00');

  s.getCell('C3').numFmt = '0%';
  assert.equal(range.numFmt, undefined, 'a block whose cells disagree says so');
});

test('a partly-styled block reports no shared facet, because its holes render unstyled', () => {
  const s = sheet();
  s.getCell('B2').font = {bold: true};
  assert.equal(s.getRange('B2:C3').font, undefined);
  assert.deepEqual(s.getRange('B2:B2').font, {bold: true});
});

test('a shared facet compares structurally, not by identity', () => {
  const s = sheet();
  s.getCell('B2').font = {name: 'Arial', bold: true};
  // The same font stated with its keys in the other order is the same font.
  s.getCell('C2').font = {bold: true, name: 'Arial'};
  assert.deepEqual(s.getRange('B2:C2').font, {name: 'Arial', bold: true});
});

test('assigning style composes facet by facet, exactly as a cell does', () => {
  const s = sheet();
  const range = s.getRange('B2:C3');
  range.font = {name: 'Arial'};
  range.style = {alignment: {horizontal: 'center'}};

  assert.deepEqual(range.font, {name: 'Arial'}, 'the omitted facet is untouched');
  assert.deepEqual(range.alignment, {horizontal: 'center'});
  assert.deepEqual(range.style, {font: {name: 'Arial'}, alignment: {horizontal: 'center'}});
});

test('clearStyle strips every facet and leaves values alone', () => {
  const s = sheet();
  s.getCell('B2').value = 'keep me';
  const range = s.getRange('B2:C3');
  range.style = {font: {bold: true}, numFmt: '0.00', alignment: {horizontal: 'center'}};
  range.clearStyle();

  assert.deepEqual(range.style, {});
  assert.equal(s.getCell('B2').value, 'keep me');
});

test('clearing one facet touches only the cells that exist', () => {
  const s = sheet();
  s.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFEEEEEE'}};
  const range = s.getRange('B2:C3');
  range.fill = undefined;

  assert.equal(s.getCell('B2').fill, undefined);
  // There is nothing to clear on a hole, so materialising the block to write nothing would be
  // pure cost — and would leave three empty cells behind that the sheet did not have.
  assert.equal(range.cells.length, 1);
});

test('styling a block that overlaps a merge restyles the master, stranding nothing on a covered cell', () => {
  const s = sheet();
  s.getCell('B2').value = 'merged';
  s.mergeCells('B2:C3');
  const range = s.getRange('B2:C3');
  range.font = {bold: true};

  assert.deepEqual(s.getCell('B2').font, {bold: true});
  // Every address in the block resolved to the master, so no covered position gained a cell of its
  // own — a style stranded there would have to be dropped on write.
  assert.equal(s.hasCell(2, 3), false);
  assert.equal(s.hasCell(3, 3), false);
});

test('a later per-cell edit inside the block overrides just that cell', () => {
  const s = sheet();
  const range = s.getRange('B2:C3');
  range.font = {name: 'Arial'};
  s.getCell('C3').font = {name: 'Courier New'};

  assert.deepEqual(s.getCell('B2').font, {name: 'Arial'});
  assert.deepEqual(s.getCell('C3').font, {name: 'Courier New'});
});
