import assert from 'node:assert/strict';
import test from 'node:test';
import {Worksheet} from './worksheet.ts';

test('a column handle is a live view — two handles on one index agree', () => {
  const sheet = new Worksheet('S', 1);
  const first = sheet.getColumn(2);
  const second = sheet.getColumn(2);

  first.width = 18;
  assert.equal(second.width, 18);

  second.getCell(1).value = 'b';
  assert.equal(first.cells.length, 1);
});

test('reading a column creates nothing — no record, no cells, no used range', () => {
  const sheet = new Worksheet('S', 1);
  const column = sheet.getColumn(50);

  assert.equal(column.properties, undefined);
  assert.deepEqual(column.cells, []);
  assert.equal(sheet.columnCount, 0, 'merely asking about column 50 does not widen the used range');
  assert.deepEqual([...sheet.columns()], []);
});

test('column.letter names the index the way a spreadsheet does', () => {
  const sheet = new Worksheet('S', 1);
  assert.equal(sheet.getColumn(1).letter, 'A');
  assert.equal(sheet.getColumn(27).letter, 'AA');
});

test('every ColumnProperties facet round-trips through the handle, style facets included', () => {
  const sheet = new Worksheet('S', 1);
  const column = sheet.getColumn(3);
  const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}} as const;
  const font = {name: 'Calibri', size: 11} as const;
  const border = {top: {style: 'thin'}} as const;
  const alignment = {horizontal: 'center'} as const;
  const protection = {locked: false} as const;

  column.key = 'name';
  column.width = 24;
  column.hidden = true;
  column.outlineLevel = 1;
  column.collapsed = true;
  column.fill = fill;
  column.numFmt = '0.00';
  column.font = font;
  column.border = border;
  column.alignment = alignment;
  column.protection = protection;

  assert.deepEqual(sheet.getColumn(3).properties, {
    key: 'name',
    width: 24,
    hidden: true,
    outlineLevel: 1,
    collapsed: true,
    fill,
    numFmt: '0.00',
    font,
    border,
    alignment,
    protection,
  });
});

test("a column's key still resolves a keyed row, now that it is set through the handle", () => {
  const sheet = new Worksheet('S', 1);
  sheet.getColumn(1).key = 'name';
  sheet.getColumn(2).key = 'joined';

  sheet.addRow({joined: 2026, name: 'Ada'});

  assert.equal(sheet.getCell('A1').value, 'Ada');
  assert.equal(sheet.getCell('B1').value, 2026);
});

test('column.cells is the materialised cells in ascending row order, sparsely', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('B3').value = 'three';
  sheet.getCell('B1').value = 'one';
  sheet.getCell('A2').value = 'other column';

  assert.deepEqual(
    sheet.getColumn(2).cells.map((cell) => cell.address),
    ['B1', 'B3'],
  );
});

test('column.values reads by position with row 1 at index 0, and assigns like a row does', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getColumn(1).values = ['a', 'b', 'c'];
  assert.deepEqual(sheet.getColumn(1).values, ['a', 'b', 'c']);
  assert.equal(sheet.getCell('A3').value, 'c');

  sheet.getColumn(1).values = [undefined, 'B'];
  assert.deepEqual(
    sheet.getColumn(1).values,
    ['a', 'B', 'c'],
    'holes skip, and a shorter array does not clear the tail',
  );
});

test('column.getCell resolves through merges, as the sheet does', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'master';
  sheet.mergeCells('A1:B2');
  assert.equal(sheet.getColumn(2).getCell(1).address, 'A1');
});

test('a column is out of bounds below 1, and rejects a non-integer', () => {
  const sheet = new Worksheet('S', 1);
  assert.throws(() => sheet.getColumn(0), RangeError);
  assert.throws(() => sheet.getColumn(1.5), RangeError);
});

test('sheet.columns() yields handles for the formatted columns, ascending', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getColumn(3).width = 12;
  sheet.getColumn(1).hidden = true;

  const columns = [...sheet.columns()];
  assert.deepEqual(
    columns.map((column) => column.index),
    [1, 3],
  );
  assert.equal(columns[1]?.width, 12);
});
