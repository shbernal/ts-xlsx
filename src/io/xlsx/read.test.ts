import assert from 'node:assert/strict';
import {test} from 'node:test';

import {zipSync, strToU8} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {isFormulaValue} from '../../core/value.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

/** Write a workbook and read it straight back — the round-trip under test. */
function roundtrip(workbook: Workbook): Workbook {
  return readXlsx(writeXlsx(workbook));
}

test('scalar cell values survive the round-trip with their types', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 42;
  sheet.getCell('A2').value = 'hello';
  sheet.getCell('A3').value = true;
  sheet.getCell('A4').value = false;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, 42);
  assert.equal(back?.getCell('A2').value, 'hello');
  assert.equal(back?.getCell('A3').value, true);
  assert.equal(back?.getCell('A4').value, false);
});

test('a string with markup-significant and leading/trailing space round-trips exactly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = '  <a> & "b" \t end  ';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, '  <a> & "b" \t end  ');
});

test('a formula with a numeric result round-trips as {formula, result}', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'SUM(B1:B2)', result: 7};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'SUM(B1:B2)');
  assert.equal(value.result, 7);
});

test('a formula with a string result carries its t="str" result back', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'CONCAT("a","b")', result: 'ab'};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.result, 'ab');
});

test('a formula with no cached result round-trips without inventing one', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'NOW()'};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.result, undefined);
});

test('column width and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).width = 24;
  sheet.getColumn(4).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(2).width, 24);
  assert.equal(back?.getColumn(4).hidden, true);
});

test('row height and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getRow(1).height = 33;
  sheet.getRow(2).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getRow(1).height, 33);
  assert.equal(back?.getRow(2).hidden, true);
});

test('merged ranges round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.mergeCells('A1:B2');
  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual([...(back?.merges ?? [])], ['A1:B2']);
});

test('page margins round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.pageMargins.left = 0.5;
  sheet.pageMargins.top = 1.25;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageMargins.left, 0.5);
  assert.equal(back?.pageMargins.top, 1.25);
});

test('workbook core properties round-trip', () => {
  const wb = new Workbook();
  wb.properties.creator = 'Ada Lovelace';
  wb.properties.lastModifiedBy = 'Grace Hopper';
  wb.properties.created = new Date('2020-01-02T03:04:05.000Z');
  wb.properties.modified = new Date('2021-06-07T08:09:10.000Z');
  wb.addWorksheet('S').getCell('A1').value = 1;

  const back = roundtrip(wb);
  assert.equal(back.properties.creator, 'Ada Lovelace');
  assert.equal(back.properties.lastModifiedBy, 'Grace Hopper');
  assert.equal(back.properties.created?.toISOString(), '2020-01-02T03:04:05.000Z');
  assert.equal(back.properties.modified?.toISOString(), '2021-06-07T08:09:10.000Z');
});

test('multiple sheets round-trip in order and are addressable by name', () => {
  const wb = new Workbook();
  for (const name of ['First', 'Second', 'Third']) wb.addWorksheet(name).getCell('A1').value = name;
  const back = roundtrip(wb);
  assert.deepEqual(
    back.worksheets.map(s => s.name),
    ['First', 'Second', 'Third']
  );
  assert.equal(back.getWorksheet('Second')?.getCell('A1').value, 'Second');
});

test('rowCount spans a gap; actualRowCount counts only populated rows', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A3').value = 'c';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.rowCount, 3);
  assert.equal(back?.actualRowCount, 2);
});

test('the inflate bound rejects a part whose declared size is over the cap', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const buffer = writeXlsx(wb);
  assert.throws(() => readXlsx(buffer, {maxUncompressedBytes: 16}), /possible zip bomb/);
});

test('a zip that is not an xlsx (no workbook part) is rejected, not misread', () => {
  const bogus = zipSync({'hello.txt': strToU8('not a spreadsheet')});
  assert.throws(() => readXlsx(bogus), /xl\/workbook\.xml is missing/);
});

test('a t="s" shared-string cell resolves against the shared table', () => {
  // Our writer emits inlineStr, but the reader must also resolve shared strings that
  // foreign generators use. Assemble a minimal package by hand to exercise that path.
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/sharedStrings.xml': strToU8('<?xml version="1.0"?><sst><si><t>shared</t></si></sst>'),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>'
    ),
  };
  const back = readXlsx(zipSync(files)).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, 'shared');
});
