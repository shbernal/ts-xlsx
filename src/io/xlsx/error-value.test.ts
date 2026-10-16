import assert from 'node:assert/strict';
import {test} from 'node:test';

import {encodeAddress} from '../../core/address.ts';
import {ERROR_CODES} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {
  assertRelationshipsWired,
  optionalPartIn,
  partIn,
  partsOf,
  partsWritten,
  patchParts,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {readSheetRows} from './read-rows.ts';
import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';
import {writeXlsx} from './write.ts';

test('an error Excel has no literal for is written as #VALUE! beside the rich value naming it', () => {
  // Excel 16.0 (build 20326) opened a cell `<c t="e"><v>…</v></c>` clean for `#N/A`, `#GETTING_DATA` and
  // `#BUSY!`, and with its repair prompt for `#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!` and
  // `#UNKNOWN!`, each of which it stores as `#VALUE!` with a `vm`.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = {error: '#SPILL!'};
  sheet.getCell('B1').value = {error: '#N/A'};
  sheet.getCell('C1').value = {formula: 'NA()', result: {error: '#CALC!'}};
  sheet.getCell('D1').value = {error: '#SPILL!'};

  const bytes = writeXlsx(wb);
  const parts = partsOf(bytes);
  const xml = partIn(parts, SHEET1);
  assert.match(xml, /<c r="A1" t="e" vm="1"><v>#VALUE!<\/v><\/c>/);
  assert.match(xml, /<c r="B1" t="e"><v>#N\/A<\/v><\/c>/, 'a literal points nowhere');
  assert.match(xml, /<c r="C1" t="e" vm="2"><f>NA\(\)<\/f><v>#VALUE!<\/v><\/c>/);
  assert.match(xml, /<c r="D1" t="e" vm="1">/, 'the same error shares its rich value');
  assert.match(partIn(parts, 'xl/richData/rdrichvalue.xml'), /<rv s="0"><v>8<\/v><v>0<\/v><\/rv>/);
  assert.match(partIn(parts, '[Content_Types].xml'), /rdrichvaluestructure\+xml/);
  assertRelationshipsWired(bytes);
});

test('every error round-trips, the ones stored as rich values included', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  ERROR_CODES.forEach((code, i) => {
    sheet.getCell(encodeAddress(i + 1, 1)).value = {error: code};
    sheet.getCell(encodeAddress(i + 1, 2)).value = {formula: 'NA()', result: {error: code}};
  });
  const bytes = writeXlsx(wb);
  const back = readXlsx(bytes).getWorksheet('S');
  const streamed = [...readSheetRows(bytes)];
  ERROR_CODES.forEach((code, i) => {
    assert.deepEqual(
      back?.getCell(encodeAddress(i + 1, 1)).value,
      {error: code},
      `${code} as a value`,
    );
    assert.deepEqual(
      back?.getCell(encodeAddress(i + 1, 2)).value,
      {formula: 'NA()', result: {error: code}},
      `${code} as a cached result`,
    );
    assert.deepEqual(streamed[0]?.cells[i]?.value, {error: code}, `${code} streamed`);
  });
});

test('a workbook with no rich-value error writes no rich-value part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#DIV/0!'};
  const parts = partsWritten(wb);
  assert.equal(optionalPartIn(parts, 'xl/metadata.xml'), undefined);
  assert.equal(optionalPartIn(parts, 'xl/richData/rdrichvalue.xml'), undefined);
});

test('the streaming writer shares the rich values across the rows it flushes', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow([{error: '#CALC!'}]).commit();
  sheet.addRow([{error: '#FIELD!'}, {error: '#CALC!'}]).commit();
  const bytes = await writer.commit();
  assert.ok(bytes !== undefined);
  const back = readXlsx(bytes).getWorksheet('S');
  assert.deepEqual(
    [back?.getCell('A1').value, back?.getCell('A2').value, back?.getCell('B2').value],
    [{error: '#CALC!'}, {error: '#FIELD!'}, {error: '#CALC!'}],
  );
  assert.match(sheetXml(bytes), /<c r="B2" t="e" vm="1">/);
});

test('a vm naming a rich-value error wins over the #VALUE! beside it', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#SPILL!'};
  const cell = '<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>';
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml.replace(
        cell,
        `${cell}<c r="B1" t="e"><v>#VALUE!</v></c><c r="C1" t="e" vm="7"><v>#VALUE!</v></c>`,
      ),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').value, {error: '#SPILL!'});
  assert.deepEqual(back?.getCell('B1').value, {error: '#VALUE!'}, 'no vm');
  assert.deepEqual(back?.getCell('C1').value, {error: '#VALUE!'}, 'a vm naming no block');
});
