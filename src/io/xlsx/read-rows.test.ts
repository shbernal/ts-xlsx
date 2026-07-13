import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8, zipSync} from 'fflate';

import {isFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {readSheetRows, type StreamedRow} from './read-rows.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function rows(data: Uint8Array, options?: Parameters<typeof readSheetRows>[1]): StreamedRow[] {
  return [...readSheetRows(data, options)];
}

test('yields rows in order, non-empty cells only, with decoded values', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'name';
  sheet.getCell('B1').value = 'qty';
  sheet.getCell('A2').value = 'widget';
  sheet.getCell('B2').value = 42;
  sheet.getCell('B3').value = true;

  const streamed = rows(writeXlsx(wb));
  assert.deepEqual(
    streamed.map((row) => row.number),
    [1, 2, 3]
  );
  assert.deepEqual(
    streamed[0]?.cells.map((cell) => [cell.address, cell.value]),
    [
      ['A1', 'name'],
      ['B1', 'qty'],
    ]
  );
  assert.deepEqual(
    streamed[1]?.cells.map((cell) => [cell.col, cell.value]),
    [
      [1, 'widget'],
      [2, 42],
    ]
  );
  // Row 3 has only B3 — A3 was never written, so it is absent, not a null cell.
  assert.deepEqual(
    streamed[2]?.cells.map((cell) => [cell.address, cell.value]),
    [['B3', true]]
  );
});

test('falsy-but-present values survive; only a truly empty cell drops', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 0;
  sheet.getCell('B1').value = false;
  sheet.getCell('C1').value = '';
  // D1 given a style but no value — a blank the data read should omit.
  sheet.getCell('D1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};

  const [row] = rows(writeXlsx(wb));
  assert.deepEqual(
    row?.cells.map((cell) => [cell.address, cell.value]),
    [
      ['A1', 0],
      ['B1', false],
      ['C1', ''],
    ]
  );
});

test('row numbers are preserved across a gap', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'top';
  sheet.getCell('A5').value = 'bottom';

  assert.deepEqual(
    rows(writeXlsx(wb)).map((row) => [row.number, row.cells[0]?.value]),
    [
      [1, 'top'],
      [5, 'bottom'],
    ]
  );
});

test('a formula cell yields a formula value with its cached result', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'SUM(B1:B2)', result: 7};

  const value = rows(writeXlsx(wb))[0]?.cells[0]?.value;
  assert.ok(value !== undefined);
  assert.ok(isFormulaValue(value));
  assert.equal(value.formula, 'SUM(B1:B2)');
  assert.equal(value.result, 7);
});

test('a modern function is unmangled back to its plain name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'XLOOKUP(B1,C:C,D:D)', result: 0};

  const value = rows(writeXlsx(wb))[0]?.cells[0]?.value;
  assert.ok(value !== undefined);
  assert.ok(isFormulaValue(value));
  assert.equal(value.formula, 'XLOOKUP(B1,C:C,D:D)');
});

test('a date value round-trips as a Date, not a bare serial', () => {
  const when = new Date(Date.UTC(2021, 5, 15));
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = when;

  const value = rows(writeXlsx(wb))[0]?.cells[0]?.value;
  assert.ok(value instanceof Date);
  assert.equal(value.getTime(), when.getTime());
});

test('selects a worksheet by name', () => {
  const wb = new Workbook();
  wb.addWorksheet('first').getCell('A1').value = 'one';
  wb.addWorksheet('second').getCell('A1').value = 'two';

  assert.equal(rows(writeXlsx(wb), {sheet: 'second'})[0]?.cells[0]?.value, 'two');
});

test('selects a worksheet by 1-based position; default is the first', () => {
  const wb = new Workbook();
  wb.addWorksheet('first').getCell('A1').value = 'one';
  wb.addWorksheet('second').getCell('A1').value = 'two';
  const data = writeXlsx(wb);

  assert.equal(rows(data)[0]?.cells[0]?.value, 'one');
  assert.equal(rows(data, {sheet: 1})[0]?.cells[0]?.value, 'one');
  assert.equal(rows(data, {sheet: 2})[0]?.cells[0]?.value, 'two');
});

test('a missing sheet selector is an error, not silent emptiness', () => {
  const wb = new Workbook();
  wb.addWorksheet('only').getCell('A1').value = 'x';
  const data = writeXlsx(wb);

  assert.throws(() => rows(data, {sheet: 'nope'}), /no worksheet named/);
  assert.throws(() => rows(data, {sheet: 3}), /no worksheet at position 3/);
});

test('an inline string cell decodes through the streaming SAX path', () => {
  // writeXlsx pools strings into sharedStrings, so hand-build an inlineStr sheet to prove the
  // `<is><t>` path the streaming reader must also handle.
  const sheetXml =
    '<?xml version="1.0"?><worksheet><sheetData>' +
    '<row r="1"><c r="A1" t="inlineStr"><is><t>hi there</t></is></c></row>' +
    '</sheetData></worksheet>';
  const archive = zipSync({
    '[Content_Types].xml': strToU8('<Types/>'),
    'xl/workbook.xml': strToU8('<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(sheetXml),
  });

  assert.equal(rows(archive)[0]?.cells[0]?.value, 'hi there');
});

test('streamed values agree with readXlsx cell-for-cell', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'text';
  sheet.getCell('B1').value = 3.14;
  sheet.getCell('A2').value = false;
  sheet.getCell('B2').value = {formula: 'A1&"!"', result: 'text!'};
  sheet.getCell('A3').value = new Date(Date.UTC(2000, 0, 1));
  const data = writeXlsx(wb);

  const model = readXlsx(data).getWorksheet('S');
  for (const row of rows(data)) {
    for (const cell of row.cells) {
      assert.deepEqual(cell.value, model?.getCell(cell.address).value, cell.address);
    }
  }
});

test('the generator is lazy: the first row is available before the rest are pulled', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let r = 1; r <= 1000; r += 1) sheet.getCell(`A${r}`).value = r;

  const iterator = readSheetRows(writeXlsx(wb));
  const first = iterator.next();
  assert.equal(first.done, false);
  assert.equal(first.value.number, 1);
  assert.equal(first.value.cells[0]?.value, 1);
  // Pulling one row must not have required draining the sheet; the next pull continues in order.
  assert.equal(iterator.next().value?.number, 2);
});
