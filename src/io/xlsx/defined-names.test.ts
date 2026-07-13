import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function workbookXmlOf(workbook: Workbook): string {
  return strFromU8(unzipSync(writeXlsx(workbook))['xl/workbook.xml'] as Uint8Array);
}

test('a global defined name is written into <definedNames> after <sheets>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'TaxRate', refersTo: 'S!$A$1'});
  const xml = workbookXmlOf(wb);
  assert.match(xml, /<definedNames><definedName name="TaxRate">S!\$A\$1<\/definedName><\/definedNames>/);
  assert.ok(xml.indexOf('<sheets>') < xml.indexOf('<definedNames>'), 'definedNames follows sheets');
});

test('no <definedNames> element is emitted when the workbook has none', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  assert.doesNotMatch(workbookXmlOf(wb), /<definedNames/);
});

test('a global defined name round-trips through write then read', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'TaxRate', refersTo: 'S!$A$1:$B$2'});
  const back = readXlsx(writeXlsx(wb));
  assert.deepEqual([...back.definedNames], [{name: 'TaxRate', refersTo: 'S!$A$1:$B$2'}]);
});

test('a sheet-scoped name carries the 0-based localSheetId of its sheet', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').getCell('A1').value = 1;
  wb.addWorksheet('Second').getCell('A1').value = 2;
  wb.defineName({name: 'Local', refersTo: 'Second!$A$1', scope: 'Second'});
  assert.match(workbookXmlOf(wb), /<definedName name="Local" localSheetId="1">Second!\$A\$1<\/definedName>/);
});

test('a sheet-scoped name round-trips back to its scope worksheet name', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').getCell('A1').value = 1;
  wb.addWorksheet('Second').getCell('A1').value = 2;
  wb.defineName({name: 'Local', refersTo: 'Second!$A$1', scope: 'Second'});
  const back = readXlsx(writeXlsx(wb));
  assert.deepEqual([...back.definedNames], [{name: 'Local', refersTo: 'Second!$A$1', scope: 'Second'}]);
});

test('a comment and the hidden flag survive the round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'Secret', refersTo: 'S!$A$1', comment: 'internal use', hidden: true});
  const back = readXlsx(writeXlsx(wb));
  assert.deepEqual([...back.definedNames], [
    {name: 'Secret', refersTo: 'S!$A$1', comment: 'internal use', hidden: true},
  ]);
});

test('special characters in the name and formula are escaped and round-trip verbatim', () => {
  const wb = new Workbook();
  wb.addWorksheet("O'Brien & Co").getCell('A1').value = 1;
  wb.defineName({name: 'Ampersand_Name', refersTo: "'O''Brien & Co'!$A$1"});
  const xml = workbookXmlOf(wb);
  assert.match(xml, /&amp;/);
  assert.doesNotMatch(xml, /&(?!(amp|lt|gt|quot|apos);)/);
  const back = readXlsx(writeXlsx(wb));
  assert.equal(back.definedNames[0]?.refersTo, "'O''Brien & Co'!$A$1");
});

test('a name defined as a modern function is stored under _xlfn. but modelled plain', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'Double', refersTo: 'LAMBDA(x,x*2)'});
  const xml = workbookXmlOf(wb);
  assert.match(xml, /<definedName name="Double">_xlfn\.LAMBDA\(x,x\*2\)<\/definedName>/);
  const back = readXlsx(writeXlsx(wb));
  assert.deepEqual([...back.definedNames], [{name: 'Double', refersTo: 'LAMBDA(x,x*2)'}]);
});

test('a plain reference carries no _xlfn. prefix and reads back verbatim', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'Region', refersTo: 'S!$A$1:$B$2'});
  assert.doesNotMatch(workbookXmlOf(wb), /_xlfn\./);
  const back = readXlsx(writeXlsx(wb));
  assert.equal(back.definedNames[0]?.refersTo, 'S!$A$1:$B$2');
});

test('a foreign name stored with _xlfn. reads back to its plain function name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'Pick', refersTo: 'XLOOKUP(1,S!$A:$A,S!$B:$B)'});
  const files = unzipSync(writeXlsx(wb));
  assert.match(strFromU8(files['xl/workbook.xml'] as Uint8Array), /_xlfn\.XLOOKUP/);
  const back = readXlsx(zipSync(files));
  assert.equal(back.definedNames[0]?.refersTo, 'XLOOKUP(1,S!$A:$A,S!$B:$B)');
});

test('defineName rejects an empty name and an unknown scope', () => {
  const wb = new Workbook();
  wb.addWorksheet('S');
  assert.throws(() => wb.defineName({name: '', refersTo: 'S!$A$1'}), /cannot be empty/);
  assert.throws(() => wb.defineName({name: 'X', refersTo: 'S!$A$1', scope: 'Nope'}), /unknown worksheet/);
});

test('a localSheetId pointing past the loaded sheets reads back as a global name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  wb.defineName({name: 'N', refersTo: 'S!$A$1'});
  const files = unzipSync(writeXlsx(wb));
  files['xl/workbook.xml'] = strToU8(
    strFromU8(files['xl/workbook.xml'] as Uint8Array).replace(
      '<definedName name="N">',
      '<definedName name="N" localSheetId="9">',
    ),
  );
  const back = readXlsx(zipSync(files));
  assert.deepEqual([...back.definedNames], [{name: 'N', refersTo: 'S!$A$1'}]);
});
