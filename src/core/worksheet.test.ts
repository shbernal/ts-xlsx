import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Worksheet} from './worksheet.ts';

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
