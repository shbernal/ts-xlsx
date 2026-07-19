import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

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

// An x14 worksheet part built like Excel's: the extended list validation lives entirely in
// `<extLst>`, with a cross-sheet source in `<xm:f>` and the target range in `<xm:sqref>`.
function sheetWithExtendedValidation(sqref: string, source: string): string {
  return (
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData/>' +
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" ' +
    'xr:uid="{A951C8AD-DA7D-45C7-ADAA-9BC7A5FE4CC1}">' +
    `<x14:formula1><xm:f>${source}</xm:f></x14:formula1><xm:sqref>${sqref}</xm:sqref>` +
    '</x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>'
  );
}

// Read an xlsx built from a hand-authored sheet1 part, so a case can feed the reader an x14 form the
// writer itself only produces on round-trip.
function readSheetPart(sheet1Xml: string): Workbook {
  const template = writeXlsx(withSheet(), {});
  const files = unzipSync(template);
  files['xl/worksheets/sheet1.xml'] = strToU8(sheet1Xml);
  return readXlsx(zipSync(files));
}

function withSheet(): Workbook {
  const workbook = new Workbook();
  workbook.addWorksheet('S');
  return workbook;
}

test('an extended (x14) list validation is read onto its cell with the cross-sheet source intact', () => {
  const workbook = readSheetPart(sheetWithExtendedValidation('A1:A1048576', 'Sheet2!$A:$A'));
  const dv = workbook.getWorksheet('S')?.dataValidationAt('A5');
  assert.ok(dv, 'a cell inside the extended range carries the validation');
  assert.equal(dv.type, 'list');
  assert.deepEqual(dv.formulae, ['Sheet2!$A:$A'], 'the foreign-sheet source survives verbatim');
});

test('an extended validation is re-serialised back into the x14 extLst block, not the standard element', () => {
  const workbook = readSheetPart(sheetWithExtendedValidation('B2:B16', 'Dropdown!$D$4:$D$8'));
  const xml = sheetXml(writeXlsx(workbook));

  assert.match(xml, /<x14:dataValidation type="list"/, 'the rule writes back to the extended form');
  assert.match(xml, /<xm:f>Dropdown!\$D\$4:\$D\$8<\/xm:f>/);
  assert.match(xml, /<xm:sqref>B2:B16<\/xm:sqref>/);
  // It must NOT be downgraded to a standard <dataValidation> (which cannot carry the cross-sheet ref).
  assert.doesNotMatch(xml, /<dataValidations\b/);
});

test('a sheet mixing a standard and an extended validation round-trips both to their own forms', () => {
  const part =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData/>' +
    '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A2">' +
    '<formula1>&quot;one,two&quot;</formula1></dataValidation></dataValidations>' +
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Sheet2!$D$3:$D$5</xm:f>' +
    '</x14:formula1><xm:sqref>A1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>' +
    '</worksheet>';
  const xml = sheetXml(writeXlsx(readSheetPart(part)));

  assert.match(
    xml,
    /<dataValidation type="list" allowBlank="1" sqref="A2">/,
    'standard rule stays standard',
  );
  assert.match(xml, /<x14:dataValidation type="list"/, 'extended rule stays extended');
  assert.match(xml, /<xm:f>Sheet2!\$D\$3:\$D\$5<\/xm:f>/);
});

test('an extended typed validation round-trips both operands through <xm:f> wrappers', () => {
  const workbook = new Workbook();
  workbook
    .addWorksheet('S')
    .addDataValidation(
      'A1:A9',
      {type: 'whole', operator: 'between', formulae: [1, 9]},
      {extended: true},
    );
  const pkg = writeXlsx(workbook);
  const xml = sheetXml(pkg);
  assert.match(xml, /<x14:formula1><xm:f>1<\/xm:f><\/x14:formula1>/);
  assert.match(xml, /<x14:formula2><xm:f>9<\/xm:f><\/x14:formula2>/);

  const dv = readXlsx(pkg).getWorksheet('S')?.dataValidationAt('A5');
  assert.deepEqual(dv?.formulae, [1, 9], 'both numeric operands survive the extended round-trip');
});
