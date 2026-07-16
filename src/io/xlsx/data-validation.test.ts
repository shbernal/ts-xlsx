import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function sheetXml(pkg: Uint8Array): string {
  return strFromU8(unzipSync(pkg)['xl/worksheets/sheet1.xml'] ?? new Uint8Array());
}

test('a list validation over a whole column writes exactly one dataValidation with the range as sqref', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('B2:B1048576', {
    type: 'list',
    allowBlank: true,
    formulae: ['"a,b,c"'],
  });
  const xml = sheetXml(writeXlsx(workbook));

  const entries = [...xml.matchAll(/<dataValidation[ >]/g)];
  assert.equal(entries.length, 1, 'one entry, not one per covered cell');
  assert.match(xml, /<dataValidation type="list" allowBlank="1" sqref="B2:B1048576">/);
  assert.match(xml, /<formula1>"a,b,c"<\/formula1>/);
});

test('a validation formula supplied with a leading = is serialised without it', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('A1', {
    type: 'list',
    formulae: ['=$AA$1:$AA$2'],
  });
  const xml = sheetXml(writeXlsx(workbook));

  assert.match(xml, /<formula1>\$AA\$1:\$AA\$2<\/formula1>/);
  assert.doesNotMatch(xml, /<formula1>=/);
});

test('a validation-free sheet emits no <dataValidations> element', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 1;
  assert.doesNotMatch(sheetXml(writeXlsx(workbook)), /dataValidation/);
});

test('a typed validation round-trips its type, operator, and numeric bounds on every covered cell', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('A1:A3', {
    type: 'whole',
    operator: 'between',
    allowBlank: true,
    formulae: [0, 9],
  });
  const reread = readXlsx(writeXlsx(workbook)).getWorksheet('S');
  assert.ok(reread);

  for (const ref of ['A1', 'A2', 'A3']) {
    const dv = reread.dataValidationAt(ref);
    assert.ok(dv, `${ref} carries the validation`);
    assert.equal(dv.type, 'whole');
    assert.equal(dv.operator, 'between');
    assert.deepEqual(dv.formulae, [0, 9], 'numeric literals read back as numbers');
  }
});

test('a numeric-typed validation whose operand is a reference keeps the reference, not NaN', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('A1', {
    type: 'whole',
    operator: 'greaterThan',
    formulae: ['L26'],
  });
  const dv = readXlsx(writeXlsx(workbook)).getWorksheet('S')?.dataValidationAt('A1');
  assert.deepEqual(dv?.formulae, ['L26'], 'the cell reference survives as a string');
});

test('a typed rule authored without an operator reads back as the default "between"', () => {
  // Excel omits operator="between" from the XML because it is the default; the reader restores it.
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('A1', {type: 'whole', formulae: [0, 9]});
  const dv = readXlsx(writeXlsx(workbook)).getWorksheet('S')?.dataValidationAt('A1');
  assert.equal(dv?.operator, 'between');
});

test('a list validation round-trips its string source verbatim', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('B1', {type: 'list', formulae: ['myNames']});
  const dv = readXlsx(writeXlsx(workbook)).getWorksheet('S')?.dataValidationAt('B1');
  assert.equal(dv?.type, 'list');
  assert.deepEqual(dv?.formulae, ['myNames'], 'the defined-name source is not coerced to a number');
});

test('validation formula text with markup-significant characters round-trips exactly', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addDataValidation('A1', {
    type: 'custom',
    formulae: ['AND(A1<10,A1>0)'],
  });
  const pkg = writeXlsx(workbook);
  assert.match(sheetXml(pkg), /<formula1>AND\(A1&lt;10,A1&gt;0\)<\/formula1>/);
  const dv = readXlsx(pkg).getWorksheet('S')?.dataValidationAt('A1');
  assert.deepEqual(dv?.formulae, ['AND(A1<10,A1>0)']);
});
