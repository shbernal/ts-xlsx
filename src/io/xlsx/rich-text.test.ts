import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {isRichTextValue, type RichTextValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function sheetXmlOf(data: Uint8Array): string {
  return strFromU8(unzipSync(data)['xl/worksheets/sheet1.xml'] as Uint8Array);
}

function richTextOf(workbook: Workbook, sheet: string, ref: string): RichTextValue {
  const value = workbook.getWorksheet(sheet)?.getCell(ref).value;
  assert.ok(value !== undefined && value !== null && isRichTextValue(value), 'expected a rich-text value');
  return value;
}

test('a rich-text cell round-trips its runs and per-run fonts', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {
    richText: [{text: 'bold', font: {bold: true}}, {text: ' plain'}],
  };

  const back = richTextOf(readXlsx(writeXlsx(wb)), 'S', 'A1');
  assert.equal(back.richText.length, 2);
  assert.equal(back.richText[0]?.text, 'bold');
  assert.equal(back.richText[0]?.font?.bold, true);
  assert.equal(back.richText[1]?.text, ' plain');
  assert.equal(back.richText[1]?.font, undefined, 'an unformatted run carries no font');
});

test('the run face name (rFont) round-trips', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {richText: [{text: 'x', font: {name: 'Arial', size: 14}}]};

  const sheetXml = sheetXmlOf(writeXlsx(wb));
  assert.match(sheetXml, /<rPr>.*<rFont val="Arial"\/>.*<\/rPr>/, 'the run face is <rFont>, not <name>');

  const run = richTextOf(readXlsx(writeXlsx(wb)), 'S', 'A1').richText[0];
  assert.equal(run?.font?.name, 'Arial');
  assert.equal(run?.font?.size, 14);
});

test('a rich-text cell serialises as an inline string of runs', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {richText: [{text: 'a', font: {italic: true}}, {text: 'b'}]};

  const sheetXml = sheetXmlOf(writeXlsx(wb));
  assert.match(sheetXml, /<c r="A1" t="inlineStr"><is><r><rPr><i\/><\/rPr><t>a<\/t><\/r><r><t>b<\/t><\/r><\/is><\/c>/);
});

test('an empty-text run is dropped, never emitted as an empty <t> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {
    richText: [{text: 'a', font: {bold: true}}, {text: '', font: {italic: true}}, {text: 'b'}],
  };

  const data = writeXlsx(wb);
  const sheetXml = sheetXmlOf(data);
  assert.doesNotMatch(sheetXml, /<t\/>|<t><\/t>/, 'no zero-length <t> element');

  const back = richTextOf(readXlsx(data), 'S', 'A1');
  assert.deepEqual(
    back.richText.map(r => r.text),
    ['a', 'b'],
    'only the two non-empty runs survive'
  );
  assert.equal(back.richText[0]?.font?.bold, true, 'the surrounding runs keep their formatting');
});

test('a formatted leading run keeps its formatting, identically to a non-leading run', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = {richText: [{text: 'here', font: {underline: true}}, {text: ' plain'}]};
  sheet.getCell('A2').value = {richText: [{text: 'plain ', font: {}}, {text: 'here', font: {underline: true}}]};

  const back = readXlsx(writeXlsx(wb));
  const lead = richTextOf(back, 'S', 'A1').richText.find(r => r.text === 'here');
  const tail = richTextOf(back, 'S', 'A2').richText.find(r => r.text === 'here');
  assert.equal(lead?.font?.underline, true, 'the underlined leading run survives');
  assert.equal(tail?.font?.underline, true, 'the same formatting survives on a non-leading run');
});

test('a run whose font is an empty object reads back with no font', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {richText: [{text: 'x', font: {}}]};

  const run = richTextOf(readXlsx(writeXlsx(wb)), 'S', 'A1').richText[0];
  assert.equal(run?.text, 'x');
  assert.equal(run?.font, undefined, 'a font with no facets is not materialised');
});

test('a rich-text hyperlink label round-trips as rich text with its target', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {
    hyperlink: 'https://example.org',
    text: {richText: [{text: 'bold', font: {bold: true}}, {text: 'plain'}]},
  };

  const cell = readXlsx(writeXlsx(wb)).getWorksheet('S')?.getCell('A1').value;
  assert.ok(cell !== undefined && cell !== null && typeof cell === 'object' && 'hyperlink' in cell);
  assert.equal(cell.hyperlink, 'https://example.org', 'the target survives');
  assert.ok(isRichTextValue(cell.text), 'the display label is rich text, not flattened');
  assert.equal(cell.text.richText[0]?.text, 'bold');
  assert.equal(cell.text.richText[0]?.font?.bold, true);
  assert.equal(cell.text.richText[1]?.text, 'plain');
});

test('rich text with markup-significant characters and edge whitespace round-trips exactly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {
    richText: [{text: ' a<b>&', font: {bold: true}}, {text: 'c ', font: {}}],
  };

  const back = richTextOf(readXlsx(writeXlsx(wb)), 'S', 'A1');
  assert.equal(back.richText[0]?.text, ' a<b>&', 'escaped characters and the leading space survive');
  assert.equal(back.richText[1]?.text, 'c ', 'the trailing space survives');
});
