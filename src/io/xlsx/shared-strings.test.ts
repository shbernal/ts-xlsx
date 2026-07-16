import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// Decode one package part back to its XML text; absent parts return undefined so a test can assert
// a part was NOT written.
function partText(pkg: Uint8Array, name: string): string | undefined {
  const bytes = unzipSync(pkg)[name];
  return bytes === undefined ? undefined : strFromU8(bytes);
}

function bookWithStrings(...values: string[]): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  values.forEach((value, i) => (sheet.getCell(`A${i + 1}`).value = value));
  return workbook;
}

test('useSharedStrings writes a sharedStrings part and stores the cell as a t="s" reference', () => {
  const pkg = writeXlsx(bookWithStrings('hello'), {useSharedStrings: true});

  assert.match(partText(pkg, 'xl/sharedStrings.xml') ?? '', /<si><t>hello<\/t><\/si>/);
  const sheet = partText(pkg, 'xl/worksheets/sheet1.xml') ?? '';
  assert.match(sheet, /t="s"><v>0<\/v>/);
  assert.doesNotMatch(sheet, /inlineStr/);
});

test('without the option, strings stay inline and no sharedStrings part is written', () => {
  const pkg = writeXlsx(bookWithStrings('hello'));

  assert.equal(partText(pkg, 'xl/sharedStrings.xml'), undefined);
  assert.match(partText(pkg, 'xl/worksheets/sheet1.xml') ?? '', /t="inlineStr"><is><t>hello<\/t>/);
});

test('an enabled workbook with no string cells never fabricates an empty sharedStrings part', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 42;
  const pkg = writeXlsx(workbook, {useSharedStrings: true});

  assert.equal(partText(pkg, 'xl/sharedStrings.xml'), undefined);
  // The workbook rels must not dangle a reference to a part that was omitted.
  assert.doesNotMatch(partText(pkg, 'xl/_rels/workbook.xml.rels') ?? '', /sharedStrings/);
});

test('an identical string is pooled once — count counts references, uniqueCount counts entries', () => {
  const pkg = writeXlsx(bookWithStrings('dup', 'dup', 'other'), {useSharedStrings: true});
  const sst = partText(pkg, 'xl/sharedStrings.xml') ?? '';

  assert.match(sst, /count="3"/);
  assert.match(sst, /uniqueCount="2"/);
  // Both duplicate cells reference index 0; the distinct one gets index 1.
  const sheet = partText(pkg, 'xl/worksheets/sheet1.xml') ?? '';
  assert.match(sheet, /r="A1"[^>]* t="s"><v>0<\/v>/);
  assert.match(sheet, /r="A2"[^>]* t="s"><v>0<\/v>/);
  assert.match(sheet, /r="A3"[^>]* t="s"><v>1<\/v>/);
});

test('shared and inline storage both read back to the same values', () => {
  const shared = readXlsx(writeXlsx(bookWithStrings('a', 'b', 'a'), {useSharedStrings: true}));
  const inline = readXlsx(writeXlsx(bookWithStrings('a', 'b', 'a')));
  for (const wb of [shared, inline]) {
    const sheet = wb.getWorksheet('S');
    assert.ok(sheet);
    assert.equal(sheet.getCell('A1').value, 'a');
    assert.equal(sheet.getCell('A2').value, 'b');
    assert.equal(sheet.getCell('A3').value, 'a');
  }
});

test('rich-text values stay inline even under the option, preserving run formatting', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = {
    richText: [{text: 'bold', font: {bold: true}}, {text: 'plain'}],
  };
  const pkg = writeXlsx(workbook, {useSharedStrings: true});

  // A rich cell is never pooled, so with no plain strings the part is absent entirely.
  assert.equal(partText(pkg, 'xl/sharedStrings.xml'), undefined);
  const value = readXlsx(pkg).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && typeof value === 'object' && 'richText' in value);
  assert.deepEqual(value.richText, [{text: 'bold', font: {bold: true}}, {text: 'plain'}]);
});
