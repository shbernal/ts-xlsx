import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {isDataTableFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function sheetXmlOf(data: Uint8Array): string {
  return strFromU8(unzipSync(data)['xl/worksheets/sheet1.xml'] as Uint8Array);
}

test('a data-table formula writes its t="dataTable" declaration with input cells', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {shareType: 'dataTable', ref: 'B2:B5', dataTableRow: true, r1: 'A1', result: 99};

  const cell = sheetXmlOf(writeXlsx(wb)).match(/<c r="B2"[\s\S]*?<\/c>/)?.[0] ?? '';
  assert.match(cell, /<f t="dataTable"/, 'the formula is emitted as the data-table kind');
  assert.match(cell, /ref="B2:B5"/, 'the data-table range is emitted');
  assert.match(cell, /dtr="1"/, 'the row-input flag is emitted');
  assert.match(cell, /r1="A1"/, 'the input cell is emitted');
  assert.match(cell, /<v>99<\/v>/, 'the cached result travels with the cell');
});

test('a data-table formula round-trips its kind, range, inputs, and result', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {shareType: 'dataTable', ref: 'B2:B5', dataTableRow: true, r1: 'A1', result: 99};

  const value = readXlsx(writeXlsx(wb)).getWorksheet('S')?.getCell('B2').value ?? null;
  assert.ok(isDataTableFormulaValue(value), 'the cell reads back as a data-table formula');
  assert.strictEqual(value.ref, 'B2:B5');
  assert.strictEqual(value.dataTableRow, true);
  assert.strictEqual(value.r1, 'A1');
  assert.strictEqual(value.result, 99);
});

test('a re-written data-table formula still declares t="dataTable" after a read-modify-write', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {shareType: 'dataTable', ref: 'B2:B5', dataTableRow: true, r1: 'A1', result: 99};

  const reloaded = readXlsx(writeXlsx(wb));
  const sheet2 = reloaded.getWorksheet('S');
  assert.ok(sheet2, 'the sheet reloads');
  sheet2.getCell('A1').value = 'edited elsewhere';
  assert.match(sheetXmlOf(writeXlsx(reloaded)), /t="dataTable"/, 'the kind survives an unrelated edit');
});
