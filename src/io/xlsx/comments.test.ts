import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function roundtrip(workbook: Workbook): Workbook {
  return readXlsx(writeXlsx(workbook));
}

test('a cell note survives the write/read round-trip', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'data';
  ws.getCell('A1').note = 'a helpful note';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A1').value, 'data');
  assert.strictEqual(back?.getCell('A1').note, 'a helpful note');
});

test('a note attaches to an otherwise-empty cell', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('B2').note = 'standalone';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('B2').value, null);
  assert.strictEqual(back?.getCell('B2').note, 'standalone');
});

test('a note stays on its own cell and does not bleed onto neighbours', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'noted';
  ws.getCell('A1').note = 'only here';
  ws.getCell('A2').value = 'plain';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A1').note, 'only here');
  assert.strictEqual(back?.getCell('A2').note, undefined);
});

test('note text with markup-significant characters round-trips exactly', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').note = '5 < 6 & "quoted"';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A1').note, '5 < 6 & "quoted"');
});

test('a note follows its cell when a row is inserted above it', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A2').value = 'body';
  ws.getCell('A2').note = 'travels';
  ws.insertRow(1, ['new header']);
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A3').value, 'body');
  assert.strictEqual(back?.getCell('A3').note, 'travels');
  assert.strictEqual(back?.getCell('A2').note, undefined);
});

test('a noted workbook emits a comments part, a VML drawing, and a legacyDrawing reference', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').note = 'x';
  const files = unzipSync(writeXlsx(wb));
  assert.ok(files['xl/comments1.xml'], 'a comments part is written');
  assert.ok(files['xl/drawings/vmlDrawing1.vml'], 'a VML drawing companion is written');
  const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.match(sheetXml, /<legacyDrawing r:id="[^"]+"\/>/);
  const contentTypes = strFromU8(files['[Content_Types].xml'] as Uint8Array);
  assert.match(contentTypes, /Extension="vml"/);
  assert.match(contentTypes, /PartName="\/xl\/comments1\.xml"/);
});

test('a note-free workbook writes no comment or VML parts', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'plain';
  const files = unzipSync(writeXlsx(wb));
  const names = Object.keys(files);
  assert.ok(!names.some(n => /comments\d+\.xml$/.test(n)));
  assert.ok(!names.some(n => /\.vml$/.test(n)));
  assert.ok(!strFromU8(files['[Content_Types].xml'] as Uint8Array).includes('Extension="vml"'));
});
