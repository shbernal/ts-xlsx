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
