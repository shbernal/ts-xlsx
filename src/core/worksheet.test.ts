import assert from 'node:assert/strict';
import {test} from 'node:test';

import {type CellValue, isSharedFormulaValue} from './value.ts';
import {Worksheet} from './worksheet.ts';

// The master address a shared-formula clone at `ref` currently points at.
function masterOf(sheet: Worksheet, ref: string): string {
  const value = sheet.getCell(ref).value;
  assert.ok(isSharedFormulaValue(value), `${ref} is a shared-formula clone`);
  return value.sharedFormula;
}

test('addressing a covered cell resolves to the merged region master', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  // Every covered address — including the master itself — returns the one master cell.
  const master = sheet.getCell('A1');
  assert.strictEqual(sheet.getCell('B2'), master);
  assert.strictEqual(sheet.getCell('A2'), master);
  assert.strictEqual(sheet.getCell('B1'), master);
});

test('a value written through a slave address lands on the master and reads back on either', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  sheet.getCell('B2').value = 'slave-write';
  assert.equal(sheet.getCell('A1').value, 'slave-write');
  assert.equal(sheet.getCell('B2').value, 'slave-write');
  // The slave position never materialises its own cell — only the master exists.
  assert.equal(sheet.hasCell(2, 2), false);
  assert.equal(sheet.hasCell(1, 1), true);
});

test('a cell outside every merged region is addressed literally', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  const outside = sheet.getCell('C3');
  assert.notStrictEqual(outside, sheet.getCell('A1'));
  assert.equal(outside.row, 3);
  assert.equal(outside.col, 3);
});

test('resolution consults merges at access time, not just at merge time', () => {
  const sheet = new Worksheet('S', 1);
  // Address the slave before any merge exists — it is its own cell.
  const before = sheet.getCell('B2');
  assert.equal(before.row, 2);
  assert.equal(before.col, 2);
  // Once the region is declared, the same address resolves to the master instead.
  sheet.mergeCells('A1:B2');
  assert.strictEqual(sheet.getCell('B2'), sheet.getCell('A1'));
});

test('an unbounded whole-column merge is declared but swallows no addressing', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A:A');
  assert.deepEqual([...sheet.merges], ['A:A']);
  // With no bounded rectangle there is no master to redirect to — A5 stays itself.
  const cell = sheet.getCell('A5');
  assert.equal(cell.row, 5);
  assert.equal(cell.col, 1);
});

test('merging a range that overlaps an existing merged region is rejected', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  // B2:C3 shares the corner cell B2 with A1:B2.
  assert.throws(() => sheet.mergeCells('B2:C3'), /overlaps/);
  // The rejected range never enters the merge list — only the first merge stands.
  assert.deepEqual([...sheet.merges], ['A1:B2']);
});

test('a range fully containing an existing merge is rejected', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('B2:C3');
  assert.throws(() => sheet.mergeCells('A1:D4'), /overlaps/);
});

test('unmergeCells removes a merge and frees its rectangle for a new one', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  assert.equal(sheet.unmergeCells('A1:B2'), true, 'the existing merge is removed');
  assert.deepEqual([...sheet.merges], [], 'the merge list is empty again');
  // A cell the merge had masked addresses independently, and the freed region re-merges.
  assert.equal(sheet.getCell('B2').row, 2);
  sheet.mergeCells('B2:C3');
  assert.deepEqual([...sheet.merges], ['B2:C3']);
});

test('unmergeCells returns false for a range that was never merged', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  assert.equal(sheet.unmergeCells('D4:E5'), false);
  assert.deepEqual([...sheet.merges], ['A1:B2'], 'no merge is removed');
});

test('a sheet carries no autofilter until one is set', () => {
  const sheet = new Worksheet('S', 1);
  assert.equal(sheet.autoFilter, undefined);
});

test('a bare range string sets a criteria-free filter, normalised to canonical corner order', () => {
  const sheet = new Worksheet('S', 1);
  sheet.autoFilter = 'C10:A1';
  assert.deepEqual(sheet.autoFilter, {ref: 'A1:C10', columns: []});
});

test('a values filter and a custom filter both round-trip through the setter, colIds validated', () => {
  const sheet = new Worksheet('S', 1);
  sheet.autoFilter = {
    ref: 'A1:C10',
    columns: [
      {colId: 0, criteria: {kind: 'values', values: ['apple', 'pear'], blank: false}},
      {
        colId: 2,
        criteria: {kind: 'custom', and: false, predicates: [{operator: 'greaterThan', val: '6'}]},
      },
    ],
  };
  assert.deepEqual(sheet.autoFilter?.columns[0], {
    colId: 0,
    criteria: {kind: 'values', values: ['apple', 'pear'], blank: false},
  });
  assert.equal(sheet.autoFilter?.columns[1]?.colId, 2);
});

test('a colId outside the filter range is rejected — a column must lie within it', () => {
  const sheet = new Worksheet('S', 1);
  assert.throws(
    () =>
      (sheet.autoFilter = {
        ref: 'A1:C10',
        columns: [{colId: 3, criteria: {kind: 'values', values: ['x'], blank: false}}],
      }),
    /colId 3 is outside/
  );
  assert.equal(sheet.autoFilter, undefined, 'the rejected filter never takes hold');
});

test('a custom filter with three predicates is rejected — Excel allows at most two', () => {
  const sheet = new Worksheet('S', 1);
  assert.throws(
    () =>
      (sheet.autoFilter = {
        ref: 'A1:A9',
        columns: [
          {
            colId: 0,
            criteria: {
              kind: 'custom',
              and: true,
              predicates: [
                {operator: 'greaterThan', val: '1'},
                {operator: 'lessThan', val: '9'},
                {operator: 'notEqual', val: '5'},
              ],
            },
          },
        ],
      }),
    /one or two predicates/
  );
});

test('clearing an autofilter with undefined removes it', () => {
  const sheet = new Worksheet('S', 1);
  sheet.autoFilter = 'A1:C10';
  sheet.autoFilter = undefined;
  assert.equal(sheet.autoFilter, undefined);
});

test('an unbounded autofilter range is rejected — a filter needs a bounded rectangle', () => {
  const sheet = new Worksheet('S', 1);
  assert.throws(() => (sheet.autoFilter = 'A:C'), /bounded rectangle/);
  assert.equal(sheet.autoFilter, undefined, 'the rejected range never takes hold');
});

test('merges that only share an edge but no cell are both allowed', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A1:B2');
  // C1:D2 abuts A1:B2 on the right without sharing a cell.
  sheet.mergeCells('C1:D2');
  assert.deepEqual([...sheet.merges], ['A1:B2', 'C1:D2']);
});

test('an unbounded merge is not overlap-checked against a bounded one', () => {
  const sheet = new Worksheet('S', 1);
  sheet.mergeCells('A:A');
  // A1:A3 geometrically sits inside column A, but the unbounded merge carries no rectangle,
  // so it participates in no overlap check — the bounded merge is accepted.
  sheet.mergeCells('A1:A3');
  assert.deepEqual([...sheet.merges], ['A:A', 'A1:A3']);
});

test('the exported model exposes the merged ranges, and assigning it reproduces them', () => {
  const src = new Worksheet('Src', 1);
  src.getCell('A1').value = 'merged';
  src.mergeCells('A1:C1');
  assert.deepEqual([...src.model.merges], ['A1:C1']);

  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;
  // The model round-trip is symmetric: whatever the getter exported, the setter reproduced.
  assert.deepEqual([...dst.merges], ['A1:C1']);
  assert.equal(dst.getCell('A1').value, 'merged');
});

test('a model round-trip carries cell values and per-cell style facets', () => {
  const src = new Worksheet('Src', 1);
  src.getCell('A1').value = 'title';
  src.getCell('B2').value = 42;
  src.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
  src.getCell('B2').font = {bold: true};

  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  assert.equal(dst.getCell('A1').value, 'title');
  assert.equal(dst.getCell('B2').value, 42);
  assert.deepEqual(dst.getCell('B2').fill, {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}});
  assert.deepEqual(dst.getCell('B2').font, {bold: true});
});

test('a model round-trip carries column, row, and page metadata', () => {
  const src = new Worksheet('Src', 1);
  src.getColumn(2).width = 18;
  src.getRow(3).height = 40;
  src.properties.defaultRowHeight = 15;
  src.pageMargins.left = 0.5;
  src.headerFooter.oddHeader = '&CReport';

  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  assert.equal(dst.getColumn(2).width, 18);
  assert.equal(dst.getRow(3).height, 40);
  assert.equal(dst.properties.defaultRowHeight, 15);
  assert.equal(dst.pageMargins.left, 0.5);
  assert.equal(dst.headerFooter.oddHeader, '&CReport');
});

test('a model round-trip carries tables and sheet-level protection', () => {
  const src = new Worksheet('Src', 1);
  src.addTable({name: 'T1', ref: 'A1', columns: [{name: 'Col'}], rowCount: 2});
  src.protect();

  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  assert.equal(dst.tables.length, 1);
  assert.equal(dst.tables[0]?.name, 'T1');
  assert.equal(dst.tables[0]?.ref, 'A1:A3');
  assert.notEqual(dst.protection, undefined);
});

test('a model round-trip carries the autofilter, criteria and all', () => {
  const src = new Worksheet('Src', 1);
  src.autoFilter = {
    ref: 'A1:C10',
    columns: [{colId: 1, criteria: {kind: 'values', values: ['West'], blank: false}}],
  };

  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  assert.deepEqual(dst.autoFilter, src.autoFilter);
});

test('assigning a model clears an autofilter the destination held, leaving no residue', () => {
  const dst = new Worksheet('Dst', 2);
  dst.autoFilter = 'Y1:Z9';

  dst.model = new Worksheet('Src', 1).model;

  assert.equal(dst.autoFilter, undefined);
});

test('assigning a model replaces content wholesale, leaving no residue', () => {
  const dst = new Worksheet('Dst', 2);
  dst.getCell('Z9').value = 'stale';
  dst.mergeCells('Y1:Z1');
  dst.pageMargins.top = 9;

  const src = new Worksheet('Src', 1);
  src.getCell('A1').value = 'fresh';

  dst.model = src.model;

  assert.equal(dst.getCell('A1').value, 'fresh');
  assert.equal(dst.hasCell(9, 26), false);
  assert.deepEqual([...dst.merges], []);
  assert.equal(dst.pageMargins.top, undefined);
});

test('the exported model does not alias the source sheet through mutable containers', () => {
  const src = new Worksheet('Src', 1);
  src.pageMargins.left = 0.25;

  const model = src.model;
  model.pageMargins.left = 99;
  // Mutating the snapshot must not reach back into the live sheet.
  assert.equal(src.pageMargins.left, 0.25);
});

test('spliceRows deletes the requested count and shifts the tail up', () => {
  const sheet = new Worksheet('S', 1);
  for (let i = 1; i <= 10; i++) sheet.getCell(`A${i}`).value = `r${i}`;
  sheet.spliceRows(3, 2);
  assert.equal(sheet.rowCount, 8);
  assert.equal(sheet.getCell('A2').value, 'r2', 'rows above the cut are untouched');
  assert.equal(sheet.getCell('A3').value, 'r5', 'r3 and r4 removed, r5 shifts up into row 3');
  assert.equal(sheet.getCell('A4').value, 'r6');
});

test('a splice count larger than the rows present clears the tail rather than doing nothing', () => {
  const sheet = new Worksheet('S', 1);
  for (let i = 1; i <= 10; i++) sheet.getCell(`A${i}`).value = `r${i}`;
  sheet.spliceRows(3, 200);
  assert.equal(sheet.rowCount, 2, 'an over-large count removes the whole tail, not zero rows');
  assert.equal(sheet.getCell('A3').value, null);
});

test('a deleted row shifts the row below up carrying its full cell style', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'top';
  const styled = sheet.getCell('A3');
  styled.value = 'styled';
  styled.font = {bold: true};
  styled.numFmt = '0.00';
  sheet.spliceRows(1, 1);
  const moved = sheet.getCell('A2');
  assert.equal(moved.value, 'styled', 'the styled cell shifts from A3 up to A2');
  assert.deepEqual(moved.font, {bold: true}, 'the font travels with the shifted cell');
  assert.equal(moved.numFmt, '0.00', 'the number format travels with the shifted cell');
});

test('inserting a row shifts the rows at and below it down', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A2').value = 'b';
  sheet.insertRow(2, ['inserted']);
  assert.equal(sheet.getCell('A1').value, 'a');
  assert.equal(sheet.getCell('A2').value, 'inserted');
  assert.equal(sheet.getCell('A3').value, 'b', 'the row formerly at A2 shifts down to A3');
});

test('addRow appends after the last used row and returns its cells', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'header';
  const cells = sheet.addRow(['a', 'b']);
  assert.equal(sheet.getCell('A2').value, 'a', 'appended below the header, not over it');
  assert.equal(sheet.getCell('B2').value, 'b');
  assert.deepEqual(
    cells.map(cell => cell.value),
    ['a', 'b'],
    'returns the materialised cells for styling',
  );
});

test('addRow lands below a formatting-only row, not over it', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getRow(3).height = 40;
  sheet.addRow(['tail']);
  assert.equal(sheet.getCell('A4').value, 'tail', 'the used range spans the formatting-only row at 3');
});

test('addRows stacks each row in order, even when value-less', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'header';
  const created = sheet.addRows([[null], ['x'], [null]]);
  assert.equal(created.length, 3);
  assert.equal(sheet.getCell('A2').value, null, 'a value-less appended row still consumes its slot');
  assert.equal(sheet.getCell('A3').value, 'x', 'the next row does not collide with the value-less one above');
  assert.equal(sheet.getCell('A4').value, null);
});

test('addRow skips a hole in a sparse array', () => {
  const sheet = new Worksheet('S', 1);
  const sparse: CellValue[] = [];
  sparse[0] = 'a';
  sparse[2] = 'c'; // index 1 stays a genuine array hole
  const cells = sheet.addRow(sparse);
  assert.equal(sheet.hasCell(1, 2), false, 'the hole leaves column B unmaterialised');
  assert.equal(sheet.getCell('C1').value, 'c');
  assert.deepEqual(
    cells.map(cell => cell.value),
    ['a', 'c'],
    'only the visited elements become cells',
  );
});

test('a row-splice shifts a merged range below the cut and keeps it merged', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'header';
  sheet.getCell('A2').value = 'banner';
  sheet.mergeCells('A2:O2');
  sheet.spliceRows(1, 1);
  assert.deepEqual([...sheet.merges], ['A1:O1'], 'deleting the row above shifts the banner up and keeps it merged');
});

test('an inserted row shifts a merged range below it down', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A2').value = 'banner';
  sheet.mergeCells('A2:C2');
  sheet.insertRow(1, ['inserted']);
  assert.ok([...sheet.merges].includes('A3:C3'), `expected A3:C3; got ${JSON.stringify([...sheet.merges])}`);
});

test('a splice far below a merged range leaves it untouched', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A2').value = 'banner';
  sheet.mergeCells('A2:O2');
  sheet.spliceRows(10, 1);
  assert.deepEqual([...sheet.merges], ['A2:O2']);
});

test('duplicateRow makes a faithful copy that carries no merge of its own', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'a';
  sheet.getCell('B1').value = 'b';
  sheet.getCell('C1').value = 'c';
  sheet.duplicateRow(1, 1, true);
  assert.equal(sheet.getCell('A2').value, 'a', 'the duplicated row copies the values');
  assert.equal(sheet.getCell('C2').value, 'c');
  // No phantom merge was fabricated on the new row, so an explicit merge succeeds.
  assert.doesNotThrow(() => sheet.mergeCells('A2:C2'));
});

test('duplicating rows above a merged range shifts the merge down by the number inserted', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A3').value = 'banner';
  sheet.mergeCells('A3:C3');
  sheet.duplicateRow(1, 2, true);
  assert.ok([...sheet.merges].includes('A5:C5'), `expected A5:C5; got ${JSON.stringify([...sheet.merges])}`);
});

test('inserting a row above a table shifts the table range down', () => {
  const sheet = new Worksheet('S', 1);
  sheet.addTable({name: 'T', ref: 'A3', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
  sheet.insertRow(1, ['inserted']);
  assert.equal(sheet.tables[0]?.ref, 'A4:B6');
});

test('getTable finds a table by name and addRow appends cells into the grid', () => {
  const sheet = new Worksheet('S', 1);
  sheet.addTable({name: 'T', ref: 'A1', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
  const table = sheet.getTable('T');
  assert.ok(table, 'the table is found by name');
  assert.equal(table.rowCount, 2);
  table.addRow(['c', 3]); // appends at row 4 (header on row 1 + 2 data rows)
  assert.equal(table.rowCount, 3);
  assert.equal(table.ref, 'A1:B4', 'the range grows to cover the appended row');
  assert.equal(sheet.getCell('A4').value, 'c', 'the appended values land in the grid');
  assert.equal(sheet.getCell('B4').value, 3);
});

test('getTable returns undefined for an unknown table name', () => {
  assert.equal(new Worksheet('S', 1).getTable('missing'), undefined);
});

test('a table column style bakes into the body cells addRow writes', () => {
  const sheet = new Worksheet('S', 1);
  const table = sheet.addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Amount', style: {numFmt: '#,##0.00'}}, {name: 'Label'}],
    rowCount: 0,
  });
  table.addRow([1234.5, 'x']);
  table.addRow([6789, 'y']);
  assert.equal(sheet.getCell('A2').numFmt, '#,##0.00', 'the styled column carries the numFmt');
  assert.equal(sheet.getCell('A3').numFmt, '#,##0.00');
  assert.equal(sheet.getCell('B2').numFmt, undefined, 'the unstyled column is unaffected');
  assert.equal(sheet.getCell('B3').numFmt, undefined);
});

test('spliceColumns removes the requested columns and shifts the rest left', () => {
  const sheet = new Worksheet('S', 1);
  ['A', 'B', 'C', 'D', 'E'].forEach((L, i) => (sheet.getCell(`${L}1`).value = `c${i + 1}`));
  sheet.spliceColumns(2, 2);
  assert.equal(sheet.columnCount, 3);
  assert.equal(sheet.getCell('A1').value, 'c1', 'columns before the cut are untouched');
  assert.equal(sheet.getCell('B1').value, 'c4', 'c2 and c3 removed, c4 shifts into column B');
  assert.equal(sheet.getCell('C1').value, 'c5');
});

test('spliceColumns can insert blank columns, shifting existing columns right', () => {
  const sheet = new Worksheet('S', 1);
  ['A', 'B', 'C', 'D', 'E'].forEach((L, i) => (sheet.getCell(`${L}1`).value = `c${i + 1}`));
  sheet.spliceColumns(3, 0, [], []);
  assert.equal(sheet.getCell('B1').value, 'c2', 'columns before the insertion point are untouched');
  assert.equal(sheet.getCell('C1').value, null, 'the inserted columns are blank');
  assert.equal(sheet.getCell('E1').value, 'c3', 'c3 shifts right by two into column E');
  assert.equal(sheet.getCell('G1').value, 'c5');
});

test('a column-splice re-anchors a merged range lying to the right of the cut', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('F1').value = 'F';
  sheet.getCell('H1').value = 'H';
  sheet.mergeCells('F1:G1');
  sheet.spliceColumns(2, 1);
  assert.ok([...sheet.merges].includes('E1:F1'), `expected E1:F1; got ${JSON.stringify([...sheet.merges])}`);
  assert.equal(sheet.getCell('G1').value, 'H', 'trailing data shifts left with the columns');
});

test('a cell note travels with its cell through a row splice', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A2').value = 'body';
  sheet.getCell('A2').note = 'travels';
  sheet.spliceRows(1, 0, ['header']);
  assert.equal(sheet.getCell('A3').value, 'body');
  assert.equal(sheet.getCell('A3').note, 'travels', 'the note follows its cell to the shifted position');
  assert.equal(sheet.getCell('A2').note, undefined, 'the inserted row carries no note');
});

test('a cell note round-trips through model export and import', () => {
  const src = new Worksheet('Src', 1);
  src.getCell('B2').value = 'v';
  src.getCell('B2').note = 'remember me';
  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;
  assert.equal(dst.getCell('B2').value, 'v');
  assert.equal(dst.getCell('B2').note, 'remember me');
});

test('inserting a column re-anchors a shared-formula clone to its master’s new address', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('B1').value = {formula: 'A1*2', result: 2};
  sheet.getCell('B2').value = {sharedFormula: 'B1', result: 4};
  sheet.getCell('B3').value = {sharedFormula: 'B1', result: 6};

  // A column inserted at 1 shifts the whole group right by one: master B1 → C1, clones B2/B3 → C2/C3.
  sheet.spliceColumns(1, 0, []);
  assert.equal(masterOf(sheet, 'C2'), 'C1', 'the clone follows its master to the new column');
  assert.equal(masterOf(sheet, 'C3'), 'C1');
});

test('inserting a row re-anchors a shared-formula clone to its master’s new address', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = {formula: 'Z1', result: 0};
  sheet.getCell('B1').value = {sharedFormula: 'A1', result: 0};
  sheet.getCell('C1').value = {sharedFormula: 'A1', result: 0};

  // A row inserted at 1 shifts the group down by one: master A1 → A2, clones B1/C1 → B2/C2.
  sheet.spliceRows(1, 0, []);
  assert.equal(masterOf(sheet, 'B2'), 'A2', 'the clone follows its master to the new row');
  assert.equal(masterOf(sheet, 'C2'), 'A2');
});

test('a splice below a shared-formula group leaves the clone’s master address untouched', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('B1').value = {formula: 'A1*2', result: 2};
  sheet.getCell('B2').value = {sharedFormula: 'B1', result: 4};

  // The master B1 sits above the edit at row 5, so nothing moves and no address is rewritten.
  sheet.spliceRows(5, 0, []);
  assert.equal(masterOf(sheet, 'B2'), 'B1', 'an untouched master keeps its original address');
});

test('addRow places a dense array left-to-right and skips holes in a sparse array', () => {
  const sheet = new Worksheet('S', 1);
  sheet.addRow(['a', 'b', 'c']);
  // eslint-disable-next-line no-sparse-arrays
  sheet.addRow(['x', , 'z']);
  assert.deepEqual(
    ['A1', 'B1', 'C1'].map(r => sheet.getCell(r).value),
    ['a', 'b', 'c']
  );
  assert.strictEqual(sheet.getCell('A2').value, 'x');
  assert.strictEqual(sheet.getCell('B2').value, null, 'the hole leaves its column untouched');
  assert.strictEqual(sheet.getCell('C2').value, 'z');
});

test('addRow places a keyed object under the columns carrying the matching key', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getColumn(1).key = 'k1';
  sheet.getColumn(2).key = 'k2';
  sheet.addRow({k1: 'o1', k2: 'o2'});
  assert.strictEqual(sheet.getCell('A1').value, 'o1');
  assert.strictEqual(sheet.getCell('B1').value, 'o2');
});

test('addRow rejects a keyed object naming a column that has no key', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getColumn(1).key = 'k1';
  assert.throws(() => sheet.addRow({missing: 1}), /no column is keyed/);
});

test('addRows populates a mixed batch of array- and object-shaped rows', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getColumn(1).key = 'k1';
  sheet.getColumn(2).key = 'k2';
  sheet.addRows([
    ['m1', 'm2'],
    {k1: 'n1'},
  ]);
  assert.deepEqual(
    ['A1', 'B1'].map(r => sheet.getCell(r).value),
    ['m1', 'm2']
  );
  assert.strictEqual(sheet.getCell('A2').value, 'n1', 'the object-shaped batch row is populated too');
});
