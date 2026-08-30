import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeRange} from './address.ts';
import {cellCarriesContent} from './cell.ts';
import {INTERNAL} from './internal.ts';
import {Worksheet} from './worksheet.ts';

// An independent recomputation of the used range: the naive scan the maintained extent replaced,
// written against the sheet's public surface. Every edit that can pull the grid inward has to leave
// the two agreeing, so this is the oracle the shrinking cases below are checked against rather than
// a hard-coded number, which would only prove the extent agrees with whoever wrote the test.
function scanned(sheet: Worksheet): {rows: number; columns: number} {
  let rows = 0;
  let columns = 0;
  for (const row of sheet.rows()) {
    if (row.properties !== undefined) rows = Math.max(rows, row.number);
    for (const cell of row.cells) {
      if (cellCarriesContent(cell)) {
        rows = Math.max(rows, cell.row);
        columns = Math.max(columns, cell.col);
      }
    }
  }
  for (const column of sheet.columns()) columns = Math.max(columns, column.index);
  for (const merge of sheet.merges) {
    const {bottom, right} = decodeRange(merge);
    rows = Math.max(rows, bottom ?? 0);
    columns = Math.max(columns, right ?? 0);
  }
  return {rows, columns};
}

function assertAgreesWithScan(sheet: Worksheet, after: string): void {
  const oracle = scanned(sheet);
  assert.equal(sheet.rowCount, oracle.rows, `rowCount disagrees with a full scan after ${after}`);
  assert.equal(
    sheet.columnCount,
    oracle.columns,
    `columnCount disagrees with a full scan after ${after}`,
  );
}

// A grid whose last row and last column are held open by their own formatting as well as by cells.
// A declared line is the case that matters here: an over-large *cell* bound costs only a scan, since
// the scan is what confirms it, but a row height or column width is taken at face value, so an edit
// that removes one has to say so or the extent keeps reporting a line that is gone.
function populated(rows = 6, columns = 4): Worksheet {
  const sheet = new Worksheet('S', 1);
  for (let row = 1; row <= rows; row++) {
    for (let col = 1; col <= columns; col++)
      sheet.getRow(row).getCell(col).value = `r${row}c${col}`;
  }
  sheet.getRow(rows).height = 30;
  sheet.getColumn(columns).width = 18;
  return sheet;
}

test('eviction leaves the extent where a full scan would put it', () => {
  const sheet = populated();
  for (const number of [6, 5, 4]) sheet[INTERNAL].evictRow(number);
  assert.equal(sheet.rowCount, 3, 'the evicted rows took their height with them');
  assertAgreesWithScan(sheet, 'evicting the tail');
});

test('assigning a model replaces the extent rather than keeping the old high-water mark', () => {
  const wide = populated(6, 4);
  const narrow = populated(2, 2);
  wide.model = narrow.model;
  assert.equal(wide.rowCount, 2);
  assert.equal(wide.columnCount, 2);
  assertAgreesWithScan(wide, 'assigning a smaller model');
});

test('a splice that deletes the tail pulls the extent in with it', () => {
  const sheet = populated();
  sheet.spliceRows(3, 4);
  assert.equal(sheet.rowCount, 2);
  assertAgreesWithScan(sheet, 'splicing the tail away');

  sheet.spliceColumns(2, 3);
  assert.equal(sheet.columnCount, 1);
  assertAgreesWithScan(sheet, 'splicing columns away');
});

test('unmerging gives back the extent the merge was holding open', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'only value';
  sheet.mergeCells('B2:E9');
  assert.equal(sheet.rowCount, 9, 'the merge occupies its whole rectangle');
  assert.equal(sheet.columnCount, 5);
  assertAgreesWithScan(sheet, 'merging');

  sheet.unmergeCells('B2:E9');
  assert.equal(sheet.rowCount, 1, 'nothing else reached past row 1');
  assert.equal(sheet.columnCount, 1);
  assertAgreesWithScan(sheet, 'unmerging');
});

test('clearing a cell the sheet never hears about still shrinks the extent', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'kept';
  // A caller holding the cell can empty it without touching the sheet, so the extent can never
  // simply trust a remembered answer: it confirms the topmost line against the live grid.
  const outlier = sheet.getCell('E9');
  outlier.value = 'temporary';
  outlier.font = {bold: true};
  assert.equal(sheet.rowCount, 9);

  outlier.value = undefined;
  outlier.font = undefined;
  assert.equal(sheet.rowCount, 1, 'the emptied cell no longer bounds the range');
  assert.equal(sheet.columnCount, 1);
  assertAgreesWithScan(sheet, 'clearing a held cell');
});

test('duplicateRow over the rows below leaves the extent agreeing with a scan', () => {
  const sheet = populated(4, 3);
  sheet.duplicateRow(1, {count: 2, insert: false});
  assertAgreesWithScan(sheet, 'duplicating without inserting');

  sheet.duplicateRow(1, {count: 2});
  assertAgreesWithScan(sheet, 'duplicating with an insert');
});

// Each appender reads the extent once per line it appends, so deriving the extent by scanning made
// appending quadratic. The guard is a ratio rather than a millisecond budget: the ratio is what tells
// linear from quadratic, and it does not depend on how fast the machine is.
//
// The timing is taken as the fastest of several runs. A single run of an allocation-heavy loop is at
// the mercy of whenever the collector decides to pause, which on this workload swings the ratio from
// 1.5x to 20x on unchanged code; the fastest run is the one that was interrupted least, and across
// repeated trials it lands within a few tenths of the true figure.
function fastestRun(runs: number, work: () => void): number {
  let best = Infinity;
  for (let i = 0; i < runs; i++) {
    const start = performance.now();
    work();
    best = Math.min(best, performance.now() - start);
  }
  return best;
}

function appendRows(count: number): void {
  const sheet = new Worksheet('S', 1);
  for (let i = 0; i < count; i++) sheet.addRow(['a', 'b', 'c', 'd', 'e']);
}

test('addRow in a loop scales linearly, not quadratically', () => {
  const small = fastestRun(5, () => appendRows(4000));
  const large = fastestRun(5, () => appendRows(16000));
  // Four times the rows. Linear measures around 4x here and the scanning extent measured 23x, so a
  // bound of 9x separates them with roughly a doubling of headroom over the worst observed run.
  assert.ok(
    large < small * 9,
    `16k rows took ${large.toFixed(1)}ms against ${small.toFixed(1)}ms for 4k: that is ${(large / small).toFixed(1)}x, which is not linear`,
  );
});
