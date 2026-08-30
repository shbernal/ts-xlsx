import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeRange} from './address.ts';
import {type Cell, cellCarriesContent} from './cell.ts';
import {INTERNAL} from './internal.ts';
import {UsedExtent} from './used-extent.ts';
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
// appending quadratic. What follows counts that work rather than timing it.
//
// It used to be timed, as a ratio between two sizes taken as the fastest of several runs. A ratio
// does tell linear from quadratic and does not depend on how fast the machine is, but it does depend
// on the machine being *free*: one `verify` run had another gate holding the CPU and the budget
// failed, then passed on three serial re-runs. A gate that fails on scheduling teaches an agent to
// re-run rather than to read. Counting the rows the extent visits is exact, is the property itself
// rather than a proxy for it, and cannot be knocked over by a neighbouring process.

/**
 * A row map that reports every row the extent looks at, whether by key or by walking. A real `Map`
 * subclass rather than an object shaped like one, so the iterator types line up exactly with what
 * `UsedExtent` declares it is handed and nothing has to be cast past the typechecker.
 */
class CountingRows extends Map<number, ReadonlyMap<number, Cell>> {
  visits = 0;

  override get(key: number): ReadonlyMap<number, Cell> | undefined {
    this.visits++;
    return super.get(key);
  }

  *#walk<T>(source: Iterable<T>): Generator<T> {
    for (const item of source) {
      this.visits++;
      yield item;
    }
  }

  override entries(): MapIterator<[number, ReadonlyMap<number, Cell>]> {
    return this.#walk(super.entries());
  }

  override keys(): MapIterator<number> {
    return this.#walk(super.keys());
  }

  override values(): MapIterator<ReadonlyMap<number, Cell>> {
    return this.#walk(super.values());
  }

  override [Symbol.iterator](): MapIterator<[number, ReadonlyMap<number, Cell>]> {
    return this.entries();
  }
}

/** An extent over `count` fully-populated rows, and the row map that counts what it visits. */
function populatedExtent(count: number): {extent: UsedExtent; rows: CountingRows} {
  const sheet = new Worksheet('S', 1);
  for (let i = 0; i < count; i++) sheet.addRow(['a', 'b', 'c', 'd', 'e']);
  return extentOver(sheet);
}

// Rebuilt from a sheet's own cells rather than from a stub, so what is counted is a walk of the real
// grid: exactly the cells the scan this class replaced would have had to visit.
function extentOver(sheet: Worksheet): {extent: UsedExtent; rows: CountingRows} {
  const rows = new CountingRows();
  for (const row of sheet.rows()) {
    rows.set(row.number, new Map(row.cells.map((cell) => [cell.col, cell])));
  }
  const extent = new UsedExtent({
    rows,
    rowProperties: new Map(),
    columns: new Map(),
    mergeRects: [],
  });
  for (const [number, cols] of new Map(rows)) {
    for (const col of cols.keys()) extent.noteCell(number, col);
  }
  rows.visits = 0;
  return {extent, rows};
}

test('reading the extent after an append visits one row, whatever the sheet already holds', () => {
  // The case the whole class exists for: an appender asks where the grid reaches, immediately after
  // putting a row at the top of it. The answer is one map lookup, and it is still one lookup at
  // sixteen times the size. The scan it replaced visited every row, which is what made appending
  // quadratic.
  for (const count of [1000, 4000, 16000]) {
    const {extent, rows} = populatedExtent(count);
    assert.equal(extent.lastRow, count, `lastRow over ${count} rows`);
    assert.equal(rows.visits, 1, `lastRow over ${count} rows visited ${rows.visits} row(s)`);
  }
});

test('an unused topmost row is what costs the walk, and it is the only thing that does', () => {
  // The bound may overstate: a row materialised by `getCell` and never filled is a key the extent
  // believes in until the read confirms it. That case falls back to the walk this class replaced,
  // and pinning it is what keeps the constant above meaningful - the fast path is a fast path, not
  // an absence of a scan, and the gap between the two counts is the whole design.
  const sheet = new Worksheet('S', 1);
  for (let i = 0; i < 1000; i++) sheet.addRow(['a']);
  assert.equal(
    cellCarriesContent(sheet.getCell('A1001')),
    false,
    'the topmost row is materialised but empty',
  );

  const {extent, rows} = extentOver(sheet);
  assert.equal(extent.lastRow, 1000, 'the empty top row does not extend the used range');
  assert.equal(
    rows.visits,
    1002,
    'the lookup that failed, then one visit per row of the fallback walk',
  );
});
