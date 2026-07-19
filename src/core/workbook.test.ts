import assert from 'node:assert/strict';
import {test} from 'node:test';

import {ValueType} from './value.ts';
import {Workbook} from './workbook.ts';

test('addWorksheet appends sheets with stable, distinct ids', () => {
  const wb = new Workbook();
  const a = wb.addWorksheet('Alpha');
  const b = wb.addWorksheet('Beta');
  assert.equal(a.id, 1);
  assert.equal(b.id, 2);
  assert.deepEqual(
    wb.worksheets.map((s) => s.name),
    ['Alpha', 'Beta'],
  );
});

test('a new worksheet defaults to visible', () => {
  const wb = new Workbook();
  assert.equal(wb.addWorksheet('S').state, 'visible');
  assert.equal(wb.addWorksheet('H', {state: 'hidden'}).state, 'hidden');
});

test('getWorksheet finds sheets by case-insensitive name and by id', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('Data');
  assert.equal(wb.getWorksheet('data'), sheet);
  assert.equal(wb.getWorksheet('DATA'), sheet);
  assert.equal(wb.getWorksheet(1), sheet);
  assert.equal(wb.getWorksheet('missing'), undefined);
  assert.equal(wb.getWorksheet(99), undefined);
});

test('duplicate sheet names are rejected case-insensitively', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(() => wb.addWorksheet('sheet1'), /already exists/);
});

test('invalid sheet names are rejected up front', () => {
  const wb = new Workbook();
  assert.throws(() => wb.addWorksheet(''), /cannot be empty/);
  assert.throws(() => wb.addWorksheet('a'.repeat(32)), /31-character limit/);
  assert.throws(() => wb.addWorksheet('a/b'), /forbids/);
  assert.throws(() => wb.addWorksheet('a:b'), /forbids/);
  assert.throws(() => wb.addWorksheet("'quoted'"), /apostrophe/);
});

test("a cell's col and row are 1-based numbers matching its position", () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const cell = sheet.getCell('B3');
  cell.value = 'x';
  assert.equal(typeof cell.col, 'number');
  assert.equal(typeof cell.row, 'number');
  assert.equal(cell.col, 2);
  assert.equal(cell.row, 3);
  assert.equal(cell.address, 'B3');
  assert.equal(cell.type, ValueType.String);
});

test('getCell returns the same cell instance on repeat access', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const first = sheet.getCell('A1');
  first.value = 7;
  assert.equal(sheet.getCell('A1'), first);
  assert.equal(sheet.getCell('$A$1').value, 7);
});

test('getCell rejects a whole-row or whole-column reference', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  assert.throws(() => sheet.getCell('1'), /not a single-cell reference/);
  assert.throws(() => sheet.getCell('A'), /not a single-cell reference/);
});

test('cells materialise lazily — only touched positions exist', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  assert.equal(sheet.hasCell(3, 2), false);
  sheet.getCell('B3');
  assert.equal(sheet.hasCell(3, 2), true);
  assert.equal(sheet.hasCell(1, 1), false);
});

test('assigning undefined clears a cell back to null/empty', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = 42;
  cell.value = undefined;
  assert.equal(cell.value, null);
  assert.equal(cell.type, ValueType.Null);
});
