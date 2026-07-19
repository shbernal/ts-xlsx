import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {dateToSerial} from '../../core/date.ts';
import {type FormulaValue, isFormulaValue, isSharedFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function sheetXmlOf(data: Uint8Array): string {
  return strFromU8(unzipSync(data)['xl/worksheets/sheet1.xml'] as Uint8Array);
}

function formulaOf(wb: Workbook, ref: string): FormulaValue {
  const value = wb.getWorksheet('S')?.getCell(ref).value;
  assert.ok(value !== undefined && value !== null && isFormulaValue(value), `${ref} is a formula`);
  return value;
}

test('a date-valued formula result round-trips as a Date, not a bare serial', () => {
  const wb = new Workbook();
  const when = new Date(2020, 0, 1);
  wb.addWorksheet('S').getCell('A1').value = {formula: 'TODAY()', result: when};

  const a1 = formulaOf(readXlsx(writeXlsx(wb)), 'A1');
  assert.equal(a1.formula, 'TODAY()');
  assert.ok(a1.result instanceof Date, 'the cached result reads back as a Date');
  assert.equal((a1.result as Date).getTime(), when.getTime());
});

test('the result caches the serial and the cell carries a date format so it reads as a date', () => {
  const wb = new Workbook();
  const when = new Date(2020, 0, 1);
  wb.addWorksheet('S').getCell('A1').value = {formula: 'TODAY()', result: when};
  const sheetXml = sheetXmlOf(writeXlsx(wb));

  // The serial rides in <v> exactly as a bare date cell stores its value, and the cell references a
  // (non-default) style — the date number format that makes the serial read back as a Date.
  assert.match(
    sheetXml,
    new RegExp(`<c r="A1" s="\\d+"><f>TODAY\\(\\)</f><v>${dateToSerial(when)}</v></c>`),
  );
});

test('an explicit date format on the cell wins over the default but still reads back as a date', () => {
  const wb = new Workbook();
  const when = new Date(2020, 5, 15);
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = {formula: 'TODAY()', result: when};
  cell.numFmt = 'yyyy-mm-dd';

  const back = readXlsx(writeXlsx(wb));
  const a1 = formulaOf(back, 'A1');
  assert.ok(a1.result instanceof Date, 'the explicit-format cell still yields a Date');
  assert.equal((a1.result as Date).getTime(), when.getTime());
  assert.equal(
    back.getWorksheet('S')?.getCell('A1').numFmt,
    'yyyy-mm-dd',
    'the chosen format survives',
  );
});

test('a formula with an Invalid Date result writes no cached value and reads back result-less', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = {formula: 'TODAY()', result: new Date(Number.NaN)};
  sheet.getCell('A2').value = 'sibling'; // one bad result must not take down the rest of the sheet

  const sheetXml = sheetXmlOf(writeXlsx(wb));
  assert.match(
    sheetXml,
    /<c r="A1"><f>TODAY\(\)<\/f><\/c>/,
    'no <v> is cached for an unrepresentable date',
  );

  const back = readXlsx(writeXlsx(wb));
  const a1 = formulaOf(back, 'A1');
  assert.equal(a1.formula, 'TODAY()');
  assert.equal(a1.result, undefined, 'no result is invented on read');
  assert.equal(back.getWorksheet('S')?.getCell('A2').value, 'sibling');
});

test('a shared-formula clone with a date result reads back as a Date', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const day1 = new Date(2021, 2, 1);
  const day2 = new Date(2021, 2, 2);
  sheet.getCell('B1').value = {formula: 'A1+1', result: day1};
  sheet.getCell('B2').value = {sharedFormula: 'B1', result: day2};

  const clone = readXlsx(writeXlsx(wb)).getWorksheet('S')?.getCell('B2').value;
  assert.ok(
    clone !== undefined && clone !== null && isSharedFormulaValue(clone),
    'B2 stays a shared formula',
  );
  assert.ok(clone.result instanceof Date, 'the clone caches its result as a Date');
  assert.equal(clone.result.getTime(), day2.getTime());
});
